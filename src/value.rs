use libffi::middle::Arg;
use windows::Win32::System::WinRT::IActivationFactory;
use windows_core::{GUID, IUnknown, Interface};
use windows_future::IAsyncInfo;

use crate::{
    WinRTType,
    call::{self, call_winrt_method_2},
    result,
};

#[derive(Debug)]
pub struct ArrayOfIUnknownData(pub windows::core::Array<IUnknown>);

impl Clone for ArrayOfIUnknownData {
    fn clone(&self) -> Self {
        let mut arr = windows::core::Array::<IUnknown>::with_len(self.0.len());
        for i in 0..self.0.len() {
            arr[i] = self.0[i].clone();
        }
        ArrayOfIUnknownData(arr)
    }
}

/// Metadata for a dynamic WinRT async operation.
#[derive(Debug, Clone)]
pub struct AsyncInfo {
    pub info: IAsyncInfo,
    pub async_type: WinRTType,
}

impl AsyncInfo {
    pub fn iid(&self) -> GUID {
        self.async_type.iid().expect("async type must have IID")
    }

    pub fn handler_iid(&self) -> GUID {
        self.async_type.completed_handler_iid().expect("async type must have handler IID")
    }

    pub fn result_type(&self) -> Option<&WinRTType> {
        match &self.async_type {
            WinRTType::IAsyncOperation(t) | WinRTType::IAsyncOperationWithProgress(t, _) => Some(t),
            _ => None,
        }
    }
}

#[derive(Debug, Clone)]
pub enum WinRTValue {
    Bool(bool),
    I8(i8),
    U8(u8),
    I16(i16),
    U16(u16),
    I32(i32),
    U32(u32),
    I64(i64),
    U64(u64),
    F32(f32),
    F64(f64),
    Object(IUnknown),
    HString(windows_core::HSTRING),
    HResult(windows_core::HRESULT),
    OutValue(*mut std::ffi::c_void, WinRTType),
    Async(AsyncInfo),
    ArrayOfIUnknown(ArrayOfIUnknownData),
}
unsafe impl Send for WinRTValue {}
unsafe impl Sync for WinRTValue {}

impl WinRTValue {
    pub fn from_activation_factory(name: &windows::core::HSTRING) -> result::Result<WinRTValue> {
        let factory = unsafe {
            windows::Win32::System::WinRT::RoGetActivationFactory::<IActivationFactory>(name)
        };
        match factory {
            Ok(factory) => Ok(WinRTValue::Object(factory.cast()?)),
            Err(e) => Err(result::Error::WindowsError(e)),
        }
    }

    pub fn as_hstring(&self) -> Option<windows::core::HSTRING> {
        match self {
            WinRTValue::HString(hstr) => Some((*hstr).clone()),
            _ => None,
        }
    }

    pub fn as_i32(&self) -> Option<i32> {
        match self {
            WinRTValue::Bool(b) => Some(*b as i32),
            WinRTValue::I8(v) => Some(*v as i32),
            WinRTValue::U8(v) => Some(*v as i32),
            WinRTValue::I16(v) => Some(*v as i32),
            WinRTValue::U16(v) => Some(*v as i32),
            WinRTValue::I32(v) => Some(*v),
            WinRTValue::U32(v) => Some(*v as i32),
            _ => None,
        }
    }

    pub fn as_object(&self) -> Option<IUnknown> {
        match self {
            WinRTValue::Object(obj) => Some(obj.clone()),
            WinRTValue::Async(a) => Some(a.info.cast().ok()?),
            _ => None,
        }
    }

    pub fn cast(&self, iid: &GUID) -> result::Result<WinRTValue> {
        match self {
            WinRTValue::Object(obj) => {
                let mut result = std::ptr::null_mut();
                unsafe { obj.query(iid, &mut result) }.ok()?;
                Ok(WinRTValue::Object(unsafe { IUnknown::from_raw(result) }))
            }
            _ => Err(result::Error::ExpectObjectTypeError(self.get_type())),
        }
    }

    pub fn call_single_out(
        &self,
        method_index: usize,
        typ: &WinRTType,
        args: &[WinRTValue],
    ) -> result::Result<WinRTValue> {
        match self {
            WinRTValue::Object(obj) => {
                let mut result = std::ptr::null_mut();
                let hr = match (typ, args) {
                    (_, []) => call::call_winrt_method_1(method_index, obj.as_raw(), &mut result),
                    (_, [WinRTValue::I32(n)]) => {
                        call_winrt_method_2(method_index, obj.as_raw(), *n, &mut result)
                    }
                    (_, [WinRTValue::I64(n)]) => {
                        call_winrt_method_2(method_index, obj.as_raw(), *n, &mut result)
                    },
                    (_, [WinRTValue::Object(x)]) => {
                        call_winrt_method_2(method_index, obj.as_raw(), x.as_raw(), &mut result)
                    }
                    _ => panic!("Unsupported number of arguments"),
                };
                hr.ok().map_err(|e| {
                    println!("Error calling method: {:?}", e);
                    result::Error::WindowsError(e)
                })?;
                Ok(typ.from_out(result).unwrap())
            }
            _ => Err(result::Error::ExpectObjectTypeError(self.get_type())),
        }
    }
    pub fn call_single_out_2(
        &self,
        method_index: usize,
        typ: &WinRTType,
        args: &[WinRTValue],
    ) -> result::Result<WinRTValue> {
        match self {
            WinRTValue::Object(obj) => {
                let mut result = typ.default_value();
                let hr = match args {
                    [] => call::call_winrt_method_1(method_index, obj.as_raw(), result.out_ptr()),
                    [WinRTValue::I32(n)] => {
                        call_winrt_method_2(method_index, obj.as_raw(), *n, result.out_ptr())
                    }
                    [WinRTValue::I64(n)] => {
                        call_winrt_method_2(method_index, obj.as_raw(), *n, result.out_ptr())
                    }
                    _ => panic!("Unsupported number of arguments"),
                };
                hr.ok().map_err(|e| result::Error::WindowsError(e))?;
                Ok(result)
            }
            _ => Err(result::Error::ExpectObjectTypeError(self.get_type())),
        }
    }
    pub fn get_type(&self) -> crate::WinRTType {
        match self {
            WinRTValue::Bool(_) => crate::WinRTType::Bool,
            WinRTValue::I8(_) => crate::WinRTType::I8,
            WinRTValue::U8(_) => crate::WinRTType::U8,
            WinRTValue::I16(_) => crate::WinRTType::I16,
            WinRTValue::U16(_) => crate::WinRTType::U16,
            WinRTValue::I32(_) => crate::WinRTType::I32,
            WinRTValue::U32(_) => crate::WinRTType::U32,
            WinRTValue::I64(_) => crate::WinRTType::I64,
            WinRTValue::U64(_) => crate::WinRTType::U64,
            WinRTValue::F32(_) => crate::WinRTType::F32,
            WinRTValue::F64(_) => crate::WinRTType::F64,
            WinRTValue::Object(_) => crate::WinRTType::Object,
            WinRTValue::HString(_) => crate::WinRTType::HString,
            WinRTValue::HResult(_) => crate::WinRTType::HResult,
            WinRTValue::OutValue(_, typ) => crate::WinRTType::OutValue(Box::new(typ.clone())),
            WinRTValue::Async(_) => crate::WinRTType::Object,
            WinRTValue::ArrayOfIUnknown(_) => crate::WinRTType::ArrayOfIUnknown,
        }
    }

    pub fn out_ptr(&mut self) -> *mut std::ffi::c_void {
        match self {
            WinRTValue::Bool(v) => v as *mut bool as _,
            WinRTValue::I8(v) => v as *mut i8 as _,
            WinRTValue::U8(v) => v as *mut u8 as _,
            WinRTValue::I16(v) => v as *mut i16 as _,
            WinRTValue::U16(v) => v as *mut u16 as _,
            WinRTValue::I32(v) => v as *mut i32 as _,
            WinRTValue::U32(v) => v as *mut u32 as _,
            WinRTValue::I64(v) => v as *mut i64 as _,
            WinRTValue::U64(v) => v as *mut u64 as _,
            WinRTValue::F32(v) => v as *mut f32 as _,
            WinRTValue::F64(v) => v as *mut f64 as _,
            WinRTValue::HString(s) => s as *mut windows_core::HSTRING as _,
            WinRTValue::Object(o) => o as *mut IUnknown as _,
            WinRTValue::HResult(hr) => hr as *mut windows_core::HRESULT as _,
            WinRTValue::OutValue(ptr, _) => *ptr,
            WinRTValue::ArrayOfIUnknown(data) => data.0.as_ptr() as *mut std::ffi::c_void,
            WinRTValue::Async(_) => panic!("Cannot get out_ptr for async value"),
        }
    }

    pub fn libffi_arg(&self) -> Arg<'_> {
        use libffi::middle::arg;
        match &self {
            WinRTValue::Bool(v) => arg(v),
            WinRTValue::I8(v) => arg(v),
            WinRTValue::U8(v) => arg(v),
            WinRTValue::I16(v) => arg(v),
            WinRTValue::U16(v) => arg(v),
            WinRTValue::I32(v) => arg(v),
            WinRTValue::U32(v) => arg(v),
            WinRTValue::I64(v) => arg(v),
            WinRTValue::U64(v) => arg(v),
            WinRTValue::F32(v) => arg(v),
            WinRTValue::F64(v) => arg(v),
            WinRTValue::Object(p) => arg(p),
            WinRTValue::HString(hstr) => arg(hstr),
            WinRTValue::HResult(hr) => arg(hr),
            WinRTValue::OutValue(p, _) => arg(p),
            WinRTValue::Async(_) => panic!("Cannot pass async value as libffi arg"),
            WinRTValue::ArrayOfIUnknown(data) => arg(&data.0),
        }
    }
}
