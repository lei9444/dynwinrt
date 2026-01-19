use core::ffi::c_void;
use windows_core::{HRESULT, HSTRING};

use crate::value;

pub fn get_vtable_function_ptr(obj: *mut c_void, method_index: usize) -> *mut c_void {
    unsafe {
        // a function pointer is *const c_void (void* in C)
        // a vtable is a array of function pointers: *const *const c_void (void** in C)
        // a COM object pointer is a pointer to vtable: *const *const *const c_void (void*** in C)
        let vtable_ptr = *(obj as *const *const *mut c_void);
        *vtable_ptr.add(method_index)
    }
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

// pub fn call_method_dynamic(
//     vtable_index: usize,
//     obj: *const c_void,
//     args: &mut [&mut value::Argument],
// ) -> HRESULT {
//     use value::Argument;
//     match args {
//         [
//             Argument {
//                 is_out: true,
//                 val: value::WinRTValue::HString(_),
//             },
//         ] => {
//             let hstr = HSTRING::new();
//             let hr = call_winrt_method_1(vtable_index, obj.cast_mut(), std::ptr::from_ref(&hstr));
//             args[0].val = value::WinRTValue::HString(hstr);
//             hr
//         }
//         _ => unimplemented!(),
//     }
// }
