use libffi::middle::Cif;
use windows::core::{GUID, HSTRING};

use crate::{call, types::WinRTType, value::WinRTValue};

#[derive(Debug, Clone)]
pub struct Parameter {
    pub typ: WinRTType,
    pub value_index: usize,
    pub is_out: bool,
}

#[derive(Debug, Clone)]
pub struct MethodSignature {
    out_count: usize,
    parameters: Vec<Parameter>,
    return_type: WinRTType,
    is_opaque: bool,
}

impl MethodSignature {
    pub fn new() -> Self {
        MethodSignature {
            out_count: 0,
            parameters: Vec::new(),
            return_type: WinRTType::HResult,
            is_opaque: false,
        }
    }

    pub fn add(mut self, typ: WinRTType) -> Self {
        self.parameters.push(Parameter {
            is_out: false,
            typ,
            value_index: self.parameters.len() - self.out_count,
        });
        self
    }

    pub fn add_out(mut self, typ: WinRTType) -> Self {
        self.parameters.push(Parameter {
            is_out: true,
            typ,
            value_index: self.out_count,
        });
        self.out_count += 1;
        self
    }

    /// Add an in-parameter of array type. At the ABI level this expands to
    /// UINT32 length + T* data.
    pub fn add_array(self, element_type: WinRTType) -> Self {
        self.add(WinRTType::Array(Box::new(element_type)))
    }

    /// Add an out-parameter of array type. At the ABI level this expands to
    /// UINT32* out_length + T** out_data (ReceiveArray pattern).
    pub fn add_out_array(self, element_type: WinRTType) -> Self {
        self.add_out(WinRTType::Array(Box::new(element_type)))
    }

    /// Add a FillArray out-parameter. Caller allocates the buffer, callee fills it.
    /// ABI: UINT32 capacity, T* items (expands to two ABI params).
    /// The actual count of filled items is typically returned as a separate U32 out-param.
    pub fn add_fill_array(self, element_type: WinRTType) -> Self {
        self.add_out(WinRTType::FillArray(Box::new(element_type)))
    }

    pub fn build(self, index: usize) -> Method {
        use libffi::middle::Type;
        let mut types: Vec<Type> = Vec::with_capacity(self.parameters.len() + 1);
        types.push(Type::pointer()); // com object's this pointer
        for param in &self.parameters {
            if param.typ.is_array() {
                if param.is_out {
                    // ReceiveArray: UINT32* out_length, T** out_data
                    types.push(Type::pointer()); // pointer to u32
                    types.push(Type::pointer()); // pointer to pointer
                } else {
                    // PassArray: UINT32 length, T* data
                    types.push(Type::u32());     // length
                    types.push(Type::pointer()); // data pointer
                }
            } else if param.typ.is_fill_array() {
                // FillArray: UINT32 capacity, T* items (caller-allocated)
                types.push(Type::u32());     // capacity
                types.push(Type::pointer()); // pointer to caller-allocated buffer
            } else if param.is_out {
                types.push(Type::pointer());
            } else {
                types.push(param.typ.libffi_type());
            }
        }
        let cif = Cif::new(types.into_iter(), self.return_type.abi_type().libffi_type());
        Method {
            info: MethodInfo {
                index,
                parameters: self.parameters,
                out_count: self.out_count,
            },
            cif,
        }
    }
}

#[derive(Debug)]
pub struct MethodInfo {
    pub index: usize,
    pub parameters: Vec<Parameter>,
    pub out_count: usize,
}

#[derive(Debug)]
pub struct Method {
    info: MethodInfo,
    cif: Cif,
}

impl Method {
    pub fn call_dynamic(
        &self,
        obj: *mut std::ffi::c_void,
        args: &[WinRTValue],
    ) -> windows_core::Result<Vec<WinRTValue>> {
        call::call_winrt_method_dynamic(
            self.info.index,
            obj,
            &self.info.parameters,
            args,
            self.info.out_count,
            &self.cif,
        )
    }
}

#[derive(Debug)]
pub struct InterfaceSignature {
    pub name: String,
    pub iid: windows_core::GUID,
    pub methods: Vec<Method>,
}

impl InterfaceSignature {
    pub fn define_interface(name: String, iid: windows_core::GUID) -> Self {
        InterfaceSignature {
            name,
            iid,
            methods: Vec::new(),
        }
    }

    pub fn define_from_iunknown(name: &str, iid: GUID) -> Self {
        let mut t = InterfaceSignature::define_interface(name.to_owned(), iid);
        t.add_method(MethodSignature::new()) // 0 QueryInterface
            .add_method(MethodSignature::new()) // 1 AddRef
            .add_method(MethodSignature::new()); // 2 Release
        t
    }

    pub fn define_from_iinspectable(name: &str, iid: GUID) -> Self {
        let mut t = Self::define_from_iunknown(name, iid);
        t.add_method(MethodSignature::new()) // 3 GetIids
            .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 4 GetRuntimeClassName
            .add_method(MethodSignature::new()); // 5 GetTrustLevel
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
