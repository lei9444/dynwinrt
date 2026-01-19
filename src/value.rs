use super::call;
use libffi::middle::{Arg, Cif, arg};
use windows_core::{HRESULT, IInspectable, IUnknown, Interface};

pub enum WinRTType {
    Object,
    HString,
    HResult,
}

// there are different stage of encodings

// Rust "managed" value - safe to use in Rust code, RAII managed lifetime of com ptrs
// C ABI encoding - lowered to C ABI types, e.g.

pub enum WinRTValue {
    Object(*mut std::ffi::c_void),
    HString(windows_core::HSTRING),
    HResult(windows_core::HRESULT),
}

pub struct Argument {
    pub is_out: bool,
    pub val: WinRTValue,
}

impl Argument {
    pub fn libffi_arg(&self) -> Arg<'_> {
        use libffi::middle::arg;
        match &self.val {
            WinRTValue::Object(p) => arg(p),
            WinRTValue::HString(hstr) => arg(hstr),
            WinRTValue::HResult(hr) => arg(hr),
        }
    }

    pub fn value_ptr(&self) -> *const std::ffi::c_void {
        match &self.val {
            WinRTValue::Object(p) => std::ptr::from_ref(p).cast(),
            WinRTValue::HString(hstr) => std::ptr::from_ref(hstr).cast(),
            WinRTValue::HResult(hr) => std::ptr::from_ref(hr).cast(),
        }
    }

    pub fn out_hstring() -> Self {
        Argument {
            is_out: true,
            val: WinRTValue::HString(Default::default()),
        }
    }

    pub fn as_hstring(&self) -> &windows_core::HSTRING {
        match &self.val {
            WinRTValue::HString(hstr) => hstr,
            _ => panic!("Not an HSTRING value"),
        }
    }
}

struct Method {
    cif: Cif,
    index: usize,
}

impl Method {
    fn new(types: &[libffi::middle::Type], index: usize) -> Self {
        let cif = Cif::new(types.into_iter().cloned(), libffi::middle::Type::i32());
        Method { cif, index }
    }
    fn get_function_pointer(&self, obj: *mut std::ffi::c_void) -> *mut std::ffi::c_void {
        call::get_vtable_function_ptr(obj, self.index)
    }
}

trait DWinRTObject {
    fn call_method(&self, method: &Method, args: &[Argument]) -> HRESULT;
}

impl<T: Interface> DWinRTObject for T {
    fn call_method(&self, method: &Method, args: &[Argument]) -> HRESULT {
        use libffi::middle::CodePtr;
        let obj = self.as_raw();
        let fptr = method.get_function_pointer(obj);
        let mut cargs: Vec<Arg> = Vec::with_capacity(args.len());
        let mut ptrs: Vec<*const std::ffi::c_void> = Vec::with_capacity(args.len());

        for i in 0..args.len() {
            if args[i].is_out {
                ptrs[i] = args[i].value_ptr();
            } else {
                cargs[i] = args[i].libffi_arg()
            }
        }
        for i in 0..args.len() {
            if args[i].is_out {
                cargs[i] = arg(&ptrs[i])
            }
        }
        unsafe { method.cif.call(CodePtr(fptr as *mut _), &cargs) }
    }
}

// WinRTType - need to describe full WinRT type system, and provide mapping to Rust types and libffi C ABI types
// WinRTValue, WinRTParameter
// ABIValue - either statically Rust direct FFI call, or dynamically via libffi
