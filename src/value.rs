use libffi::middle::Arg;
use windows_core::IUnknown;

#[derive(Debug, PartialEq, Eq)]
pub enum WinRTValue {
    I32(i32),
    Object(IUnknown),
    HString(windows_core::HSTRING),
    HResult(windows_core::HRESULT),
    Pointer(*mut std::ffi::c_void),
}

impl WinRTValue {
    pub fn as_hstring(&self) -> Option<windows_core::HSTRING> {
        match self {
            WinRTValue::HString(hstr) => Some((*hstr).clone()),
            _ => None,
        }
    }

    pub fn as_i32(&self) -> Option<i32> {
        match self {
            WinRTValue::I32(i) => Some(*i),
            _ => None,
        }
    }

    pub fn libffi_arg(&self) -> Arg<'_> {
        use libffi::middle::arg;
        match &self {
            WinRTValue::Object(p) => arg(p),
            WinRTValue::HString(hstr) => arg(hstr),
            WinRTValue::HResult(hr) => arg(hr),
            WinRTValue::I32(i) => arg(i),
            WinRTValue::Pointer(p) => arg(p),
        }
    }
}

#[derive(Debug)]
pub enum AbiValue {
    I32(i32),
    Pointer(*mut std::ffi::c_void),
}

impl AbiValue {
    pub fn out_ptr(&self) -> *const std::ffi::c_void {
        match self {
            AbiValue::I32(i) => std::ptr::from_ref(i).cast(),
            AbiValue::Pointer(p) => std::ptr::from_ref(p).cast(),
        }
    }
}
