use core::ffi::c_void;
use windows::Win32::Foundation::HSTR;
use windows_core::{GUID, HRESULT, HSTRING, IInspectable, IUnknown, Interface};

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

pub enum DWinRTValueUnion {
    Void,
    HString(windows_core::HSTRING),
    HRESULT(windows_core::HRESULT),
    Guid(GUID),
    Object(*mut c_void),
    Pointer(*mut c_void),
}

impl DWinRTValueUnion {}

pub trait DWinRTValueKind {
    type Sig;
    fn value(&self) -> Self::Sig;
    fn from_value(val: Self::Sig) -> Self;
}

pub struct DWinRTPointerValue(pub *mut c_void);
impl DWinRTPointerValue {
    pub fn from_com_object<T: Interface>(val: &T) -> Self {
        let ptr = val.as_raw() as *mut c_void;
        DWinRTPointerValue(ptr)
    }
    pub fn from_out_ptr<T>(ptr: &mut T) -> Self {
        DWinRTPointerValue(ptr as *mut _ as *mut c_void)
    }
}

impl DWinRTValueKind for DWinRTPointerValue {
    type Sig = *mut c_void;

    fn value(&self) -> Self::Sig {
        self.0
    }

    fn from_value(val: Self::Sig) -> Self {
        DWinRTPointerValue(val)
    }
}

pub struct DWinRTHRESULTValue(pub windows_core::HRESULT);
impl DWinRTValueKind for DWinRTHRESULTValue {
    type Sig = windows_core::HRESULT;

    fn value(&self) -> Self::Sig {
        self.0
    }

    fn from_value(val: Self::Sig) -> Self {
        DWinRTHRESULTValue(val)
    }
}

pub fn call_method_1_result<TI, TR>(vtable_index: usize, obj: DWinRTPointerValue, x: TI) -> TR
where
    TI: DWinRTValueKind,
    TR: DWinRTValueKind,
{
    unsafe {
        let method_ptr = get_vtable_function_ptr(obj.value(), vtable_index);

        let method: extern "system" fn(*mut c_void, TI::Sig) -> TR::Sig =
            std::mem::transmute(method_ptr);

        let result = method(obj.value(), x.value());
        TR::from_value(result)
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

pub fn call_method_1<TI: DWinRTValueKind>(
    vtable_index: usize,
    obj: DWinRTPointerValue,
    x: TI,
) -> DWinRTHRESULTValue {
    call_method_1_result::<TI, DWinRTHRESULTValue>(vtable_index, obj, x)
}

pub fn call_method_ptr_ptr_ret_hresult(
    obj: DWinRTValueUnion,
    method_ptr: *mut c_void,
    args: &[DWinRTValueUnion],
) -> DWinRTValueUnion {
    unsafe {
        type MethodType = extern "system" fn(*mut c_void, *mut c_void) -> windows_core::HRESULT;

        let method: MethodType = std::mem::transmute(method_ptr);

        let obj_ptr = match obj {
            DWinRTValueUnion::Pointer(p) => p,
            _ => panic!("Expected Pointer for object"),
        };

        let arg0_ptr = match args.get(0) {
            Some(DWinRTValueUnion::Pointer(p)) => *p,
            _ => panic!("Expected Pointer for first argument"),
        };

        let hr = method(obj_ptr, arg0_ptr);
        DWinRTValueUnion::HRESULT(hr)
    }
}

pub fn call_method_with_values(
    vtable_index: usize,
    obj: *const c_void,
    args: &mut [&mut value::Argument],
) -> HRESULT {
    use value::Argument;
    match args {
        [
            Argument {
                is_out: true,
                val: value::WinRTValue::HString(_),
            },
        ] => {
            let hstr = HSTRING::new();
            let hr = call_winrt_method_1(vtable_index, obj.cast_mut(), std::ptr::from_ref(&hstr));
            args[0].val = value::WinRTValue::HString(hstr);
            hr
        },
        _ => unimplemented!(),
    }
}
