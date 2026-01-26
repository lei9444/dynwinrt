use core::ffi::c_void;
use libffi::middle::{Arg, Cif, arg};
use windows_core::{HRESULT, HSTRING, IUnknown, Interface};

use crate::{
    signature::Parameter,
    value::{AbiValue, WinRTValue},
};

pub fn get_vtable_function_ptr(obj: *mut c_void, method_index: usize) -> *mut c_void {
    unsafe {
        // a function pointer is *const c_void (void* in C)
        // a vtable is a array of function pointers: *const *const c_void (void** in C)
        // a COM object pointer is a pointer to vtable: *const *const *const c_void (void*** in C)
        let vtable_ptr = *(obj as *const *const *mut c_void);
        *vtable_ptr.add(method_index)
    }
}

unsafe fn foo(com_this_ptr: *mut c_void, out_value: *mut *mut c_void) {}

fn usage(com_this_ptr: *mut c_void) {
    // stack allocated pointer to receive out parameter
    let mut out_value: *mut c_void = core::ptr::null_mut();
    // calling winrt methods
    unsafe {
        foo(com_this_ptr, &mut out_value);
    }
    // then I need to convert out_value to appropriate type
    let outCom : IUnknown = unsafe { IUnknown::from_raw(out_value) };
    let outHString : HSTRING = unsafe { std::mem::transmute(out_value) };
}

pub fn call_winrt_method_1<T1>(vtable_index: usize, obj: *mut c_void, x1: T1) -> HRESULT {
    let method_ptr = get_vtable_function_ptr(obj, vtable_index);

    unsafe {
        let method: extern "system" fn(*mut c_void, T1) -> HRESULT =
            std::mem::transmute(method_ptr);
        method(obj, x1)
    }
}

pub fn call_winrt_method_2<T1, T2>(
    vtable_index: usize,
    obj: *mut c_void,
    x1: T1,
    x2: T2,
) -> HRESULT {
    let method_ptr = get_vtable_function_ptr(obj, vtable_index);

    unsafe {
        let method: extern "system" fn(*mut c_void, T1, T2) -> HRESULT =
            std::mem::transmute(method_ptr);
        method(obj, x1, x2)
    }
}

pub fn call_winrt_method_dynamic(
    vtable_index: usize,
    obj: *mut c_void,
    parameters: &[Parameter],
    args: &[WinRTValue],
    out_count: usize,
    cif: &libffi::middle::Cif,
) -> windows_core::Result<Vec<WinRTValue>> {
    use libffi::middle::CodePtr;
    let fptr = get_vtable_function_ptr(obj, vtable_index);
    let mut ffi_args: Vec<Arg> = Vec::with_capacity(parameters.len() + 1);
    let mut out_values: Vec<AbiValue> = Vec::with_capacity(out_count);
    let mut out_ptrs: Vec<*const std::ffi::c_void> = Vec::with_capacity(out_count);

    ffi_args.push(arg(&obj));

    for p in parameters {
        if p.is_out {
            out_values.push(p.typ.abi_type().default_value());
            out_ptrs.push(out_values.last().unwrap().out_ptr());
        }
    }
    for p in parameters {
        if p.is_out {
            ffi_args.push(arg(&out_ptrs[p.value_index]));
        } else {
            ffi_args.push(args[p.value_index].libffi_arg());
        }
    }
    let hr: windows_core::HRESULT = unsafe { cif.call(CodePtr(fptr), &ffi_args) };
    hr.ok()?;
    let mut result_values: Vec<WinRTValue> = Vec::with_capacity(out_count);
    for p in parameters {
        if p.is_out {
            let out_value = p.typ.from_out_value(&out_values[p.value_index]);
            result_values.push(out_value);
        }
    }
    Ok(result_values)
}
