use windows_core::{IUnknown, Interface};

use crate::value::{AbiValue, WinRTValue};

#[derive(Debug)]
pub enum WinRTType {
    I32,
    Object,
    HString,
    HResult,
    Pointer,
}

impl WinRTType {
    pub fn new_out_value(&self) -> AbiValue {
        match self {
            WinRTType::I32 => AbiValue::I32(0),
            WinRTType::Object => AbiValue::Pointer(std::ptr::null_mut()),
            WinRTType::HString => AbiValue::Pointer(std::ptr::null_mut()),
            WinRTType::HResult => AbiValue::I32(0),
            WinRTType::Pointer => AbiValue::Pointer(std::ptr::null_mut()),
        }
    }

    pub fn abi_type(&self) -> AbiType {
        match self {
            WinRTType::I32 | WinRTType::HResult => AbiType::I32,
            WinRTType::Object | WinRTType::HString | WinRTType::Pointer => AbiType::Ptr,
        }
    }

    pub fn from_out_value(&self, out: &AbiValue) -> WinRTValue {
        match (self, out) {
            (WinRTType::I32, AbiValue::I32(i)) => WinRTValue::I32(*i),
            (WinRTType::Object, AbiValue::Pointer(p)) => {
                WinRTValue::Object(unsafe { IUnknown::from_raw(*p) })
            }
            (WinRTType::HString, AbiValue::Pointer(p)) => {
                WinRTValue::HString(unsafe { core::mem::transmute(*p) })
            }
            (WinRTType::HResult, AbiValue::I32(hr)) => {
                WinRTValue::HResult(windows_core::HRESULT(*hr))
            }
            (WinRTType::Pointer, AbiValue::Pointer(p)) => WinRTValue::Pointer(*p),
            _ => panic!("Mismatched out value type"),
        }
    }
}

#[derive(Debug)]
pub enum AbiType {
    I32,
    Ptr,
}

impl AbiType {
    pub fn libffi_type(&self) -> libffi::middle::Type {
        match self {
            AbiType::I32 => libffi::middle::Type::i32(),
            AbiType::Ptr => libffi::middle::Type::pointer(),
        }
    }
}
