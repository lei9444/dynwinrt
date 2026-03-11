use libffi::middle::Cif;
use std::sync::Arc;
use windows::core::{GUID, HSTRING, Interface};

use crate::{call, metadata_table::{TypeHandle, TypeKind, MetadataTable}, value::WinRTValue};

#[derive(Debug, Clone)]
pub struct Parameter {
    pub typ: TypeHandle,
    pub value_index: usize,
    pub is_out: bool,
}

#[derive(Debug, Clone)]
pub struct MethodSignature {
    out_count: usize,
    parameters: Vec<Parameter>,
    return_type: TypeHandle,
    is_opaque: bool,
    table: Arc<MetadataTable>,
}

impl MethodSignature {
    pub fn new(table: &Arc<MetadataTable>) -> Self {
        MethodSignature {
            out_count: 0,
            parameters: Vec::new(),
            return_type: table.hresult(),
            is_opaque: false,
            table: Arc::clone(table),
        }
    }

    pub fn new_with_registry(table: &Arc<MetadataTable>) -> Self {
        Self::new(table)
    }

    pub fn add_in(mut self, typ: TypeHandle) -> Self {
        self.parameters.push(Parameter {
            is_out: false,
            typ,
            value_index: self.parameters.len() - self.out_count,
        });
        self
    }

    pub fn add_out(mut self, typ: TypeHandle) -> Self {
        self.parameters.push(Parameter {
            is_out: true,
            typ,
            value_index: self.out_count,
        });
        self.out_count += 1;
        self
    }

    pub fn build(self, index: usize) -> Method {
        use libffi::middle::Type;
        let mut types: Vec<Type> = Vec::with_capacity(self.parameters.len() + 1);
        types.push(Type::pointer()); // com object's this pointer
        for param in &self.parameters {
            if param.typ.is_array() {
                if param.is_out {
                    // ReceiveArray: UINT32* out_length, T** out_data
                    types.push(Type::pointer());
                    types.push(Type::pointer());
                } else {
                    // PassArray: UINT32 length, T* data
                    types.push(Type::u32());
                    types.push(Type::pointer());
                }
            } else if param.typ.is_fill_array() {
                // FillArray: UINT32 capacity, T* items
                types.push(Type::u32());
                types.push(Type::pointer());
            } else if param.is_out {
                types.push(Type::pointer());
            } else {
                types.push(param.typ.libffi_type());
            }
        }
        let in_count = self.parameters.len() - self.out_count;
        let has_complex_param = self.parameters.iter().any(|p| {
            p.typ.is_array() || p.typ.is_fill_array() || matches!(p.typ.kind(), TypeKind::Struct(_))
        });

        // Check if the single in-param (if any) is a simple non-HString, non-Struct type
        let simple_in = !has_complex_param && in_count == 1 && {
            let in_param = self.parameters.iter().find(|p| !p.is_out).unwrap();
            !matches!(in_param.typ.kind(), TypeKind::HString)
        };

        let strategy = if !has_complex_param && in_count == 0 && self.out_count == 1 {
            CallStrategy::Direct0In1Out
        } else if !has_complex_param && in_count == 0 && self.out_count == 0 {
            CallStrategy::Direct0In0Out
        } else if simple_in && self.out_count == 0 {
            CallStrategy::Direct1In0Out
        } else if simple_in && self.out_count == 1 {
            CallStrategy::Direct1In1Out
        } else {
            CallStrategy::Libffi(Cif::new(types.into_iter(), self.return_type.abi_type().libffi_type()))
        };

        Method {
            info: MethodInfo {
                index,
                parameters: self.parameters,
                out_count: self.out_count,
            },
            strategy,
        }
    }
}

#[derive(Debug)]
pub struct MethodInfo {
    pub index: usize,
    pub parameters: Vec<Parameter>,
    pub out_count: usize,
}

/// How a Method should be invoked — decided once at build time.
#[derive(Debug)]
enum CallStrategy {
    /// 0 in + 0 out: fn(this) -> HRESULT.
    Direct0In0Out,
    /// 0 in + 1 out (getter): fn(this, out) -> HRESULT.
    Direct0In1Out,
    /// 1 in + 0 out (setter, non-HString): fn(this, val) -> HRESULT.
    Direct1In0Out,
    /// 1 in + 1 out (factory/query, non-HString in): fn(this, val, out) -> HRESULT.
    Direct1In1Out,
    /// General case → libffi via cached Cif.
    Libffi(Cif),
}

#[derive(Debug)]
pub struct Method {
    info: MethodInfo,
    strategy: CallStrategy,
}

impl Method {
    pub fn call_dynamic(
        &self,
        obj: *mut std::ffi::c_void,
        args: &[WinRTValue],
    ) -> windows_core::Result<Vec<WinRTValue>> {
        match &self.strategy {
            CallStrategy::Direct0In0Out => {
                // 0 in + 0 out: fn(this) -> HRESULT
                let hr = call::call_winrt_method_0(self.info.index, obj);
                hr.ok()?;
                Ok(vec![])
            }
            CallStrategy::Direct0In1Out => {
                // 0 in + 1 out: fn(this, out) -> HRESULT
                let param = &self.info.parameters[0];
                let mut out = param.typ.default_winrt_value();
                let hr = call::call_winrt_method_1(self.info.index, obj, out.out_ptr());
                hr.ok()?;
                // For async types, the default value is Object(null) which got filled with
                // the async COM pointer. Convert to WinRTValue::Async via from_out.
                if param.typ.is_async() {
                    let old = std::mem::replace(&mut out, WinRTValue::I32(0));
                    if let WinRTValue::Object(o) = old {
                        let ptr = o.as_raw();
                        std::mem::forget(o); // Transfer ownership to from_out
                        out = param.typ.from_out(ptr)
                            .map_err(|e| windows_core::Error::new(windows_core::HRESULT(-1), &format!("{:?}", e)))?;
                    }
                }
                Ok(vec![out])
            }
            CallStrategy::Direct1In0Out => {
                // 1 in + 0 out: fn(this, val) -> HRESULT
                let hr = call::call_1in(self.info.index, obj, &args[0]);
                hr.ok()?;
                Ok(vec![])
            }
            CallStrategy::Direct1In1Out => {
                // 1 in + 1 out: fn(this, val, out) -> HRESULT
                let out_param = self.info.parameters.iter().find(|p| p.is_out).unwrap();
                let mut out = out_param.typ.default_winrt_value();
                let hr = call::call_1in_1out(self.info.index, obj, &args[0], out.out_ptr());
                hr.ok()?;
                if out_param.typ.is_async() {
                    if let WinRTValue::Object(ref o) = out {
                        out = out_param.typ.from_out(o.as_raw())
                            .map_err(|e| windows_core::Error::new(windows_core::HRESULT(-1), &format!("{:?}", e)))?;
                    }
                }
                Ok(vec![out])
            }
            CallStrategy::Libffi(cif) => {
                call::call_winrt_method_dynamic(
                    self.info.index,
                    obj,
                    &self.info.parameters,
                    args,
                    self.info.out_count,
                    cif,
                )
            }
        }
    }
}

#[derive(Debug)]
pub struct InterfaceSignature {
    pub name: String,
    pub iid: windows_core::GUID,
    pub methods: Vec<Method>,
    table: Arc<MetadataTable>,
}

impl InterfaceSignature {
    pub fn define_interface(name: String, iid: windows_core::GUID, table: &Arc<MetadataTable>) -> Self {
        InterfaceSignature {
            name,
            iid,
            methods: Vec::new(),
            table: Arc::clone(table),
        }
    }

    pub fn define_from_iunknown(name: &str, iid: GUID, table: &Arc<MetadataTable>) -> Self {
        let mut t = InterfaceSignature::define_interface(name.to_owned(), iid, table);
        t.add_method(MethodSignature::new(table)) // 0 QueryInterface
            .add_method(MethodSignature::new(table)) // 1 AddRef
            .add_method(MethodSignature::new(table)); // 2 Release
        t
    }

    pub fn define_from_iinspectable(name: &str, iid: GUID, table: &Arc<MetadataTable>) -> Self {
        let mut t = Self::define_from_iunknown(name, iid, table);
        t.add_method(MethodSignature::new(table)) // 3 GetIids
            .add_method(MethodSignature::new(table).add_out(table.hstring())) // 4 GetRuntimeClassName
            .add_method(MethodSignature::new(table)); // 5 GetTrustLevel
        t
    }

    pub fn add_method(&mut self, signature: MethodSignature) -> &mut Self {
        let method = signature.build(self.methods.len());
        self.methods.push(method);
        self
    }
}

pub struct RuntimeClassSignature {
    name: HSTRING,
    static_interfaces: Vec<InterfaceSignature>,
    instance_interfaces: Vec<InterfaceSignature>,
}
