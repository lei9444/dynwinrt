use std::vec;

use super::call;
use libffi::middle::{Arg, Cif, arg};
use windows_core::{HRESULT, Interface};

#[derive(Debug)]
pub enum WinRTType {
    I32,
    Object,
    HString,
    HResult,
}

impl WinRTType {
    pub fn default_value(&self) -> WinRTValue {
        match self {
            WinRTType::I32 => WinRTValue::I32(0),
            WinRTType::Object => WinRTValue::Object(std::ptr::null_mut()),
            WinRTType::HString => WinRTValue::HString(windows_core::HSTRING::new()),
            WinRTType::HResult => WinRTValue::HResult(windows_core::HRESULT(0)),
        }
    }

    pub fn libffi_type(&self) -> libffi::middle::Type {
        match self {
            WinRTType::I32 => libffi::middle::Type::i32(),
            WinRTType::Object => libffi::middle::Type::pointer(),
            WinRTType::HString => libffi::middle::Type::pointer(),
            WinRTType::HResult => libffi::middle::Type::i32(),
        }
    }
}

// there are different stage of encodings

// Rust "managed" value - safe to use in Rust code, RAII managed lifetime of com ptrs
// C ABI encoding - lowered to C ABI types, e.g.

#[derive(Debug, PartialEq, Eq)]
pub enum WinRTValue {
    I32(i32),
    Object(*mut std::ffi::c_void),
    HString(windows_core::HSTRING),
    HResult(windows_core::HRESULT),
}

impl WinRTValue {
    pub fn as_hstring(&self) -> &windows_core::HSTRING {
        match self {
            WinRTValue::HString(hstr) => hstr,
            _ => panic!("Not an HSTRING value"),
        }
    }

    pub fn as_i32(&self) -> i32 {
        match self {
            WinRTValue::I32(i) => *i,
            _ => panic!("Not an i32 value"),
        }
    }

    pub fn value_ptr(&self) -> *const std::ffi::c_void {
        match &self {
            WinRTValue::Object(p) => std::ptr::from_ref(p).cast(),
            WinRTValue::HString(hstr) => std::ptr::from_ref(hstr).cast(),
            WinRTValue::HResult(hr) => std::ptr::from_ref(hr).cast(),
            WinRTValue::I32(i) => std::ptr::from_ref(i).cast(),
        }
    }

    pub fn libffi_arg(&self) -> Arg<'_> {
        use libffi::middle::arg;
        match &self {
            WinRTValue::Object(p) => arg(p),
            WinRTValue::HString(hstr) => arg(hstr),
            WinRTValue::HResult(hr) => arg(hr),
            WinRTValue::I32(i) => arg(i),
        }
    }
}

pub struct Parameter {
    typ: WinRTType,
    value_index: usize,
    is_out: bool,
}

pub struct MethodSignature {
    index: usize,
    out_count: usize,
    parameters: Vec<Parameter>,
    cif: Option<Cif>,
}

impl MethodSignature {
    pub fn new(index: usize) -> Self {
        MethodSignature {
            index,
            out_count: 0,
            parameters: Vec::new(),
            cif: None,
        }
    }

    pub fn add(&mut self, typ: WinRTType) -> &mut Self {
        self.parameters.push(Parameter {
            is_out: false,
            typ,
            value_index: self.parameters.len() - self.out_count,
        });
        self
    }

    pub fn add_out(&mut self, typ: WinRTType) -> &mut Self {
        self.parameters.push(Parameter {
            is_out: true,
            typ,
            value_index: self.out_count,
        });
        self.out_count += 1;
        self
    }

    pub fn build_cif(&mut self) {
        use libffi::middle::Type;
        let mut types: Vec<Type> = Vec::with_capacity(self.parameters.len() + 1);
        types.push(Type::pointer()); // this pointer
        for param in &self.parameters {
            if (param.is_out) {
                // out parameters are always pointers
                types.push(Type::pointer());
                continue;
            }
            types.push(match param.typ {
                WinRTType::Object => Type::pointer(),
                WinRTType::HString => Type::pointer(),
                WinRTType::HResult => Type::i32(),
                WinRTType::I32 => Type::i32(),
            });
        }
        self.cif = Some(Cif::new(types.into_iter(), Type::i32()));
    }

    pub fn call(
        &self,
        obj: *mut std::ffi::c_void,
        args: &[WinRTValue],
    ) -> windows_core::Result<Vec<WinRTValue>> {
        use libffi::middle::CodePtr;
        let fptr = call::get_vtable_function_ptr(obj, self.index);
        let mut cargs: Vec<Arg> = Vec::with_capacity(self.parameters.len() + 1);
        let mut outValues: Vec<WinRTValue> = Vec::with_capacity(self.out_count);
        let mut outPtrs: Vec<*const std::ffi::c_void> = Vec::with_capacity(self.out_count);

        cargs.push(arg(&obj));

        for p in &self.parameters {
            if p.is_out {
                outValues.push(p.typ.default_value());
                outPtrs.push(outValues.last().unwrap().value_ptr());
            }
        }
        for p in &self.parameters {
            if p.is_out {
                cargs.push(arg(&outPtrs[p.value_index]));
            } else {
                cargs.push(args[p.value_index].libffi_arg());
            }
        }
        let cif = self.cif.as_ref().unwrap();
        let hr: HRESULT = unsafe { cif.call(CodePtr(fptr), &cargs) };
        hr.ok()?;
        Ok(outValues)
    }
}

pub struct VTableSignature {
    pub methods: Vec<MethodSignature>,
}

impl VTableSignature {
    pub fn new() -> Self {
        VTableSignature {
            methods: Vec::new(),
        }
    }

    pub fn add_method(
        &mut self,
        builder: fn(sig: &mut MethodSignature) -> &MethodSignature,
    ) -> &mut Self {
        let mut method = MethodSignature::new(self.methods.len());
        builder(&mut method);
        method.build_cif();
        self.methods.push(method);
        self
    }
}

// WinRTType - need to describe full WinRT type system, and provide mapping to Rust types and libffi C ABI types
// WinRTValue, WinRTParameter
// ABIValue - either statically Rust direct FFI call, or dynamically via libffi
