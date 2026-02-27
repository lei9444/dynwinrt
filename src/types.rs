use windows_core::{GUID, IUnknown, Interface};
use windows_metadata::Signature;

use crate::abi::{AbiType, AbiValue};
use crate::signature;
use crate::value::WinRTValue;

// Well-known async interface IIDs/PIIDs (from windows_future crate)

pub const IASYNC_ACTION: GUID = windows_future::IAsyncAction::IID;
/// IAsyncActionWithProgress<P> — generic, this is the PIID
pub const IASYNC_ACTION_WITH_PROGRESS: GUID =
    GUID::from_u128(0x1f6db258_e803_48a1_9546_eb7353398884);
pub const IASYNC_OPERATION: GUID =
    GUID::from_u128(0x9fc2b0bb_e446_44e2_aa61_9cab8f636af2);
pub const IASYNC_OPERATION_WITH_PROGRESS: GUID =
    GUID::from_u128(0x9fc2b0bb_e446_44e2_aa61_9cab8f636af3);
pub const IVECTOR: GUID =
    GUID::from_u128(0x913337e9_11a1_4345_a3a2_4e7f956e222d);
pub const IVECTOR_VIEW: GUID =
    GUID::from_u128(0xbbe1fa4c_b0e3_4583_baef_1f1b2e483e56);
pub const IITERABLE: GUID =
    GUID::from_u128(0xfaa585ea_6214_4217_afda_7f46de5869b3);
pub const IITERATOR: GUID =
    GUID::from_u128(0x6a79e863_4300_459a_9966_cbb660963ee1);
pub const IMAP: GUID =
    GUID::from_u128(0x3c2925fe_8519_45c1_aa79_197b6718c1c1);
pub const IMAP_VIEW: GUID =
    GUID::from_u128(0xe9bdaaf0_cbf6_4c39_de49_316b34326a17);
pub const IKEY_VALUE_PAIR: GUID =
    GUID::from_u128(0x02b51929_c1c4_4a7e_8940_0312b5c18500);
pub const IOBSERVABLE_VECTOR: GUID =
    GUID::from_u128(0x5917eb53_50b4_4a0d_b309_65862b3f1dbc);
pub const IREFERENCE: GUID =
    GUID::from_u128(0x61c17706_2d65_11e0_9ae8_d48564015472);

// Well-known completion handler PIIDs.
// These are defined by the WinRT type system and match the values in
// windows-future's SIGNATURE strings (e.g. "pinterface({fcdcf02c-...}; ...)").
// windows-future does not export these PIIDs as standalone constants, so we
// define them here for runtime handler IID computation.

/// AsyncActionCompletedHandler — non-generic, this IS the final IID.
pub const ASYNC_ACTION_COMPLETED_HANDLER: GUID =
    windows_future::AsyncActionCompletedHandler::IID;
/// AsyncOperationCompletedHandler<T> — PIID (concrete IID computed at runtime).
pub const ASYNC_OPERATION_COMPLETED_HANDLER: GUID =
    GUID::from_u128(0xfcdcf02c_e5d8_4478_915a_4d90b74b83a5);
/// AsyncActionWithProgressCompletedHandler<P> — PIID.
pub const ASYNC_ACTION_WITH_PROGRESS_COMPLETED_HANDLER: GUID =
    GUID::from_u128(0x9c029f91_cc84_44fd_ac26_0a6c4e555281);
/// AsyncOperationWithProgressCompletedHandler<T, P> — PIID.
pub const ASYNC_OPERATION_WITH_PROGRESS_COMPLETED_HANDLER: GUID =
    GUID::from_u128(0xe85df41d_6aa7_46e3_a8e2_f009d840c627);

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum WinRTType {
    // Primitive types
    Bool,
    I8,
    U8,
    I16,
    U16,
    I32,
    U32,
    I64,
    U64,
    F32,
    F64,
    Char16,
    HString,
    Guid,

    // Composite types
    /// An untyped COM object pointer (IUnknown). Used when the concrete interface is unknown.
    Object,
    /// A non-parameterized interface, identified by its IID.
    Interface(GUID),
    /// A delegate type, identified by its IID.
    Delegate(GUID),
    /// A runtime class: `rc(FullTypeName;{default-interface-iid})`
    RuntimeClass(String, GUID),
    /// A generic (parameterized) interface definition, e.g. `IAsyncOperation<>`.
    /// Cannot be QI'd directly — must be instantiated via `Parameterized`.
    Generic { piid: GUID, arity: u32 },
    /// A parameterized interface instantiation: `pinterface({piid};arg1;arg2;...)`
    /// The first element is the generic definition (e.g., `Generic { piid, arity }`).
    Parameterized(Box<WinRTType>, Vec<WinRTType>),

    // Async patterns — sugar over Parameterized(Generic{IASYNC_*}, [...])
    IAsyncAction,
    IAsyncActionWithProgress(Box<WinRTType>),
    IAsyncOperation(Box<WinRTType>),
    IAsyncOperationWithProgress(Box<WinRTType>, Box<WinRTType>),

    // ABI-only concepts
    HResult,
    OutValue(Box<WinRTType>),
    ArrayOfIUnknown,
}

impl WinRTType {
    /// Returns true if this type is one of the four WinRT async patterns.
    pub fn is_async(&self) -> bool {
        match self {
            WinRTType::IAsyncAction
            | WinRTType::IAsyncActionWithProgress(_)
            | WinRTType::IAsyncOperation(_)
            | WinRTType::IAsyncOperationWithProgress(_, _) => true,
            WinRTType::Parameterized(generic_def, _) => Self::is_async_def(generic_def),
            _ => false,
        }
    }

    /// Produce the WinRT type signature string for this type.
    ///
    /// Used for computing parameterized interface IIDs via UUID v5.
    /// Panics for ABI-only types (HResult, OutValue, ArrayOfIUnknown) which
    /// have no WinRT type signature.
    pub fn signature_string(&self) -> std::string::String {
        match self {
            WinRTType::Bool => "b1".into(),
            WinRTType::I8 => "i1".into(),
            WinRTType::U8 => "u1".into(),
            WinRTType::I16 => "i2".into(),
            WinRTType::U16 => "u2".into(),
            WinRTType::I32 => "i4".into(),
            WinRTType::U32 => "u4".into(),
            WinRTType::I64 => "i8".into(),
            WinRTType::U64 => "u8".into(),
            WinRTType::F32 => "f4".into(),
            WinRTType::F64 => "f8".into(),
            WinRTType::Char16 => "c2".into(),
            WinRTType::HString => "string".into(),
            WinRTType::Guid => "g16".into(),
            WinRTType::Interface(iid) | WinRTType::Generic { piid: iid, .. } => format_guid_braced(iid),
            WinRTType::Delegate(iid) => {
                format!("delegate({})", format_guid_braced(iid))
            }
            WinRTType::RuntimeClass(name, default_iid) => {
                format!("rc({};{})", name, format_guid_braced(default_iid))
            }
            WinRTType::Parameterized(generic_def, args) => {
                let refs: Vec<&WinRTType> = args.iter().collect();
                pinterface_signature(&generic_def.signature_string(), &refs)
            }
            WinRTType::IAsyncAction => format_guid_braced(&IASYNC_ACTION),
            WinRTType::IAsyncActionWithProgress(p) => {
                pinterface_signature(&format_guid_braced(&IASYNC_ACTION_WITH_PROGRESS), &[p])
            }
            WinRTType::IAsyncOperation(t) => {
                pinterface_signature(&format_guid_braced(&IASYNC_OPERATION), &[t])
            }
            WinRTType::IAsyncOperationWithProgress(t, p) => {
                pinterface_signature(&format_guid_braced(&IASYNC_OPERATION_WITH_PROGRESS), &[t, p])
            }
            WinRTType::Object | WinRTType::HResult | WinRTType::OutValue(_) | WinRTType::ArrayOfIUnknown => {
                panic!("Type {:?} has no WinRT type signature", self)
            }
        }
    }

    /// Return the concrete IID for this type.
    ///
    /// For `Object`, `Delegate`, and `RuntimeClass`, returns the stored IID directly.
    /// For `Parameterized`, computes the IID on demand via UUID v5.
    /// Returns `None` for primitives and ABI-only types.
    pub fn iid(&self) -> Option<GUID> {
        match self {
            WinRTType::Interface(iid) | WinRTType::Delegate(iid) | WinRTType::RuntimeClass(_, iid) => {
                Some(*iid)
            }
            WinRTType::Parameterized(_, _) => {
                // IID is computed from the full signature string
                let sig = self.signature_string();
                let buf = windows_core::imp::ConstBuffer::from_slice(sig.as_bytes());
                Some(GUID::from_signature(buf))
            }
            WinRTType::IAsyncAction => Some(IASYNC_ACTION),
            WinRTType::IAsyncActionWithProgress(_)
            | WinRTType::IAsyncOperation(_)
            | WinRTType::IAsyncOperationWithProgress(_, _) => {
                let sig = self.signature_string();
                let buf = windows_core::imp::ConstBuffer::from_slice(sig.as_bytes());
                Some(GUID::from_signature(buf))
            }
            _ => None,
        }
    }

    /// Return the IID of the completion handler needed for `SetCompleted`.
    ///
    /// Only valid for async types. Returns `None` for non-async types.
    pub fn completed_handler_iid(&self) -> Option<GUID> {
        match self {
            WinRTType::IAsyncAction => Some(ASYNC_ACTION_COMPLETED_HANDLER),
            WinRTType::IAsyncOperation(t) => {
                let sig = pinterface_signature(
                    &format_guid_braced(&ASYNC_OPERATION_COMPLETED_HANDLER),
                    &[t],
                );
                let buf = windows_core::imp::ConstBuffer::from_slice(sig.as_bytes());
                Some(GUID::from_signature(buf))
            }
            WinRTType::IAsyncActionWithProgress(p) => {
                let sig = pinterface_signature(
                    &format_guid_braced(&ASYNC_ACTION_WITH_PROGRESS_COMPLETED_HANDLER),
                    &[p],
                );
                let buf = windows_core::imp::ConstBuffer::from_slice(sig.as_bytes());
                Some(GUID::from_signature(buf))
            }
            WinRTType::IAsyncOperationWithProgress(t, p) => {
                let sig = pinterface_signature(
                    &format_guid_braced(&ASYNC_OPERATION_WITH_PROGRESS_COMPLETED_HANDLER),
                    &[t, p],
                );
                let buf = windows_core::imp::ConstBuffer::from_slice(sig.as_bytes());
                Some(GUID::from_signature(buf))
            }
            _ => None,
        }
    }

    pub fn abi_type(&self) -> AbiType {
        match self {
            WinRTType::Bool => AbiType::Bool,
            WinRTType::I8 => AbiType::I8,
            WinRTType::U8 => AbiType::U8,
            WinRTType::I16 => AbiType::I16,
            WinRTType::U16 | WinRTType::Char16 => AbiType::U16,
            WinRTType::I32 | WinRTType::HResult => AbiType::I32,
            WinRTType::U32 => AbiType::U32,
            WinRTType::I64 => AbiType::I64,
            WinRTType::U64 => AbiType::U64,
            WinRTType::F32 => AbiType::F32,
            WinRTType::F64 => AbiType::F64,

            WinRTType::HString | WinRTType::Guid
            | WinRTType::Object | WinRTType::Interface(_) | WinRTType::Delegate(_)
            | WinRTType::RuntimeClass(_, _) | WinRTType::Parameterized(_, _)
            | WinRTType::IAsyncAction | WinRTType::IAsyncActionWithProgress(_)
            | WinRTType::IAsyncOperation(_) | WinRTType::IAsyncOperationWithProgress(_, _)
            | WinRTType::OutValue(_) | WinRTType::ArrayOfIUnknown => AbiType::Ptr,

            WinRTType::Generic { piid, .. } => {
                panic!("Cannot get ABI type for uninstantiated Generic({:?})", piid)
            }
        }
    }

    pub fn default_value(&self) -> WinRTValue {
        match self {
            WinRTType::Bool => WinRTValue::Bool(false),
            WinRTType::I8 => WinRTValue::I8(0),
            WinRTType::U8 => WinRTValue::U8(0),
            WinRTType::I16 => WinRTValue::I16(0),
            WinRTType::U16 | WinRTType::Char16 => WinRTValue::U16(0),
            WinRTType::I32 => WinRTValue::I32(0),
            WinRTType::U32 => WinRTValue::U32(0),
            WinRTType::I64 => WinRTValue::I64(0),
            WinRTType::U64 => WinRTValue::U64(0),
            WinRTType::F32 => WinRTValue::F32(0.0),
            WinRTType::F64 => WinRTValue::F64(0.0),

            WinRTType::Object | WinRTType::Interface(_) | WinRTType::Delegate(_) | WinRTType::RuntimeClass(_, _) => {
                WinRTValue::Object(unsafe { IUnknown::from_raw(std::ptr::null_mut()) })
            }

            WinRTType::HString => WinRTValue::HString(windows_core::HSTRING::new()),

            WinRTType::Guid => panic!("Cannot create default value for Guid (16-byte struct not yet supported)"),

            WinRTType::HResult => WinRTValue::HResult(windows_core::HRESULT(0)),

            WinRTType::OutValue(_) => WinRTValue::OutValue(std::ptr::null_mut(), self.clone()),

            WinRTType::Generic { piid, .. } => {
                panic!("Cannot create default value for Generic({:?})", piid)
            }

            WinRTType::Parameterized(generic_def, _) => {
                panic!("Cannot create default value for Parameterized({:?})", generic_def)
            }

            WinRTType::IAsyncAction
            | WinRTType::IAsyncActionWithProgress(_)
            | WinRTType::IAsyncOperation(_)
            | WinRTType::IAsyncOperationWithProgress(_, _) => {
                panic!("Cannot create default value for async type {:?}", self)
            }

            WinRTType::ArrayOfIUnknown => {
                WinRTValue::ArrayOfIUnknown(crate::value::ArrayOfIUnknownData(windows::core::Array::new()))
            }
        }
    }

    /// Convert a raw out pointer to a WinRTValue. Used by the fast-path call functions
    /// (`call_winrt_method_1`/`call_winrt_method_2`) which return raw pointers.
    pub fn from_out(&self, ptr: *mut std::ffi::c_void) -> crate::result::Result<WinRTValue> {
        unsafe {
            match &self {
                WinRTType::Bool => Ok(WinRTValue::Bool(*(ptr as *mut u8) != 0)),
                WinRTType::I8 => Ok(WinRTValue::I8(*(ptr as *mut i8))),
                WinRTType::U8 => Ok(WinRTValue::U8(*(ptr as *mut u8))),
                WinRTType::I16 => Ok(WinRTValue::I16(*(ptr as *mut i16))),
                WinRTType::U16 | WinRTType::Char16 => Ok(WinRTValue::U16(*(ptr as *mut u16))),
                WinRTType::I32 => Ok(WinRTValue::I32(*(ptr as *mut i32))),
                WinRTType::U32 => Ok(WinRTValue::U32(*(ptr as *mut u32))),
                WinRTType::I64 => Ok(WinRTValue::I64(*(ptr as *mut i64))),
                WinRTType::U64 => Ok(WinRTValue::U64(*(ptr as *mut u64))),
                WinRTType::F32 => Ok(WinRTValue::F32(*(ptr as *mut f32))),
                WinRTType::F64 => Ok(WinRTValue::F64(*(ptr as *mut f64))),

                WinRTType::Object | WinRTType::Interface(_) | WinRTType::Delegate(_) | WinRTType::RuntimeClass(_, _) => {
                    Ok(WinRTValue::Object(IUnknown::from_raw(ptr)))
                }

                WinRTType::HString => Ok(WinRTValue::HString(std::mem::transmute(ptr))),

                WinRTType::HResult => Ok(WinRTValue::HResult(windows_core::HRESULT(
                    *(ptr as *mut i32),
                ))),

                WinRTType::Parameterized(generic_def, args) => {
                    if Self::is_async_def(generic_def) {
                        let raw = IUnknown::from_raw(ptr);
                        let iid = self.iid().unwrap();
                        Self::make_async_value(raw, generic_def, iid, args)
                    } else {
                        Ok(WinRTValue::Object(IUnknown::from_raw(ptr)))
                    }
                }

                WinRTType::IAsyncAction
                | WinRTType::IAsyncActionWithProgress(_)
                | WinRTType::IAsyncOperation(_)
                | WinRTType::IAsyncOperationWithProgress(_, _) => {
                    let raw = IUnknown::from_raw(ptr);
                    let info: windows_future::IAsyncInfo = raw.cast()
                        .map_err(|e| crate::result::Error::WindowsError(e))?;
                    Ok(WinRTValue::Async(crate::value::AsyncInfo {
                        info,
                        async_type: self.clone(),
                    }))
                }

                _ => Err(crate::result::Error::InvalidTypeAbiToWinRT(
                    self.clone(),
                    AbiType::Ptr,
                )),
            }
        }
    }

    pub fn from_out_value(&self, out: &AbiValue) -> crate::result::Result<WinRTValue> {
        use crate::result::Error;
        match (self, out) {
            (WinRTType::Bool, AbiValue::Bool(v)) => Ok(WinRTValue::Bool(*v != 0)),
            (WinRTType::I8, AbiValue::I8(v)) => Ok(WinRTValue::I8(*v)),
            (WinRTType::U8, AbiValue::U8(v)) => Ok(WinRTValue::U8(*v)),
            (WinRTType::I16, AbiValue::I16(v)) => Ok(WinRTValue::I16(*v)),
            (WinRTType::U16 | WinRTType::Char16, AbiValue::U16(v)) => Ok(WinRTValue::U16(*v)),
            (WinRTType::I32, AbiValue::I32(v)) => Ok(WinRTValue::I32(*v)),
            (WinRTType::U32, AbiValue::U32(v)) => Ok(WinRTValue::U32(*v)),
            (WinRTType::I64, AbiValue::I64(v)) => Ok(WinRTValue::I64(*v)),
            (WinRTType::U64, AbiValue::U64(v)) => Ok(WinRTValue::U64(*v)),
            (WinRTType::F32, AbiValue::F32(v)) => Ok(WinRTValue::F32(*v)),
            (WinRTType::F64, AbiValue::F64(v)) => Ok(WinRTValue::F64(*v)),

            (WinRTType::Object | WinRTType::Interface(_) | WinRTType::Delegate(_) | WinRTType::RuntimeClass(_, _), AbiValue::Pointer(p)) => {
                Ok(WinRTValue::Object(unsafe { IUnknown::from_raw(*p) }))
            }

            (WinRTType::HString, AbiValue::Pointer(p)) => {
                Ok(WinRTValue::HString(unsafe { core::mem::transmute(*p) }))
            }

            (WinRTType::HResult, AbiValue::I32(hr)) => {
                Ok(WinRTValue::HResult(windows_core::HRESULT(*hr)))
            }

            (WinRTType::Parameterized(generic_def, args), AbiValue::Pointer(p)) => {
                if Self::is_async_def(generic_def) {
                    let raw = unsafe { IUnknown::from_raw(*p) };
                    let iid = self.iid().unwrap();
                    Self::make_async_value(raw, generic_def, iid, args)
                } else {
                    Ok(WinRTValue::Object(unsafe { IUnknown::from_raw(*p) }))
                }
            }

            (WinRTType::IAsyncAction
            | WinRTType::IAsyncActionWithProgress(_)
            | WinRTType::IAsyncOperation(_)
            | WinRTType::IAsyncOperationWithProgress(_, _), AbiValue::Pointer(p)) => {
                // Reuse from_out — AbiValue::Pointer holds the same raw ptr
                self.from_out(*p)
            }

            (WinRTType::OutValue(_), _) => Err(Error::InvalidNestedOutType(self.clone())),
            _ => Err(crate::result::Error::InvalidTypeAbiToWinRT(
                self.clone(),
                out.abi_type(),
            )),
        }
    }

    /// Check if a generic definition is one of the four WinRT async patterns.
    fn is_async_def(generic_def: &WinRTType) -> bool {
        let piid = match generic_def {
            WinRTType::Generic { piid, .. } => *piid,
            WinRTType::Interface(iid) => *iid,
            _ => return false,
        };
        piid == IASYNC_ACTION
            || piid == IASYNC_OPERATION
            || piid == IASYNC_ACTION_WITH_PROGRESS
            || piid == IASYNC_OPERATION_WITH_PROGRESS
    }

    /// Build a `WinRTValue::Async` from a raw interface pointer.
    /// `args` are the type parameters of the generic async type.
    fn make_async_value(
        raw: IUnknown,
        generic_def: &WinRTType,
        iid: GUID,
        args: &[WinRTType],
    ) -> crate::result::Result<WinRTValue> {
        let piid = match generic_def {
            WinRTType::Generic { piid, .. } => *piid,
            WinRTType::Interface(iid) => *iid,
            _ => return Err(crate::result::Error::WindowsError(
                windows_core::Error::from_hresult(windows_core::HRESULT(0x80004002u32 as i32)),
            )),
        };

        let info: windows_future::IAsyncInfo = raw.cast()
            .map_err(|e| crate::result::Error::WindowsError(e))?;

        let async_type = if piid == IASYNC_ACTION {
            WinRTType::IAsyncAction
        } else if piid == IASYNC_OPERATION {
            let t = args.first().cloned().unwrap_or(WinRTType::Object);
            WinRTType::IAsyncOperation(Box::new(t))
        } else if piid == IASYNC_ACTION_WITH_PROGRESS {
            let p = args.first().cloned().unwrap_or(WinRTType::Object);
            WinRTType::IAsyncActionWithProgress(Box::new(p))
        } else if piid == IASYNC_OPERATION_WITH_PROGRESS {
            let t = args.first().cloned().unwrap_or(WinRTType::Object);
            let p = args.get(1).cloned().unwrap_or(WinRTType::Object);
            WinRTType::IAsyncOperationWithProgress(Box::new(t), Box::new(p))
        } else {
            return Err(crate::result::Error::WindowsError(
                windows_core::Error::from_hresult(windows_core::HRESULT(0x80004002u32 as i32)),
            ));
        };

        if true {
            Ok(WinRTValue::Async(crate::value::AsyncInfo { info, async_type }))
        } else {
            Err(crate::result::Error::WindowsError(
                windows_core::Error::from_hresult(windows_core::HRESULT(0x80004002u32 as i32)),
            ))
        }
    }
}

/// Format a GUID as `{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}` (lowercase, with braces).
pub(crate) fn format_guid_braced(guid: &GUID) -> String {
    format!("{{{:?}}}", guid).to_lowercase()
}

/// Build a pinterface signature string: `pinterface({piid};arg1;arg2;...)`
fn pinterface_signature(piid_sig: &str, args: &[&WinRTType]) -> String {
    let mut s = format!("pinterface({}", piid_sig);
    for arg in args {
        s.push(';');
        s.push_str(&arg.signature_string());
    }
    s.push(')');
    s
}

/// Compute the IID of a parameterized completion handler from its PIID and type args.
fn compute_parameterized_handler_iid(handler_piid: &GUID, args: &[WinRTType]) -> GUID {
    let refs: Vec<&WinRTType> = args.iter().collect();
    let sig = pinterface_signature(&format_guid_braced(handler_piid), &refs);
    let buf = windows_core::imp::ConstBuffer::from_slice(sig.as_bytes());
    GUID::from_signature(buf)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_abi_type_mapping() {
        assert_eq!(WinRTType::Bool.abi_type(), AbiType::Bool);
        assert_eq!(WinRTType::I8.abi_type(), AbiType::I8);
        assert_eq!(WinRTType::U8.abi_type(), AbiType::U8);
        assert_eq!(WinRTType::I16.abi_type(), AbiType::I16);
        assert_eq!(WinRTType::U16.abi_type(), AbiType::U16);
        assert_eq!(WinRTType::Char16.abi_type(), AbiType::U16);
        assert_eq!(WinRTType::I32.abi_type(), AbiType::I32);
        assert_eq!(WinRTType::U32.abi_type(), AbiType::U32);
        assert_eq!(WinRTType::I64.abi_type(), AbiType::I64);
        assert_eq!(WinRTType::U64.abi_type(), AbiType::U64);
        assert_eq!(WinRTType::F32.abi_type(), AbiType::F32);
        assert_eq!(WinRTType::F64.abi_type(), AbiType::F64);
        assert_eq!(WinRTType::HResult.abi_type(), AbiType::I32);
        assert_eq!(WinRTType::HString.abi_type(), AbiType::Ptr);
        let iid = GUID::from_u128(0x12345678_1234_1234_1234_123456789abc);
        assert_eq!(WinRTType::Interface(iid).abi_type(), AbiType::Ptr);
        assert_eq!(WinRTType::RuntimeClass("Test".into(), iid).abi_type(), AbiType::Ptr);
        assert_eq!(WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IASYNC_OPERATION, arity: 1 }), vec![WinRTType::HString]).abi_type(), AbiType::Ptr);
    }

    #[test]
    fn test_signature_string() {
        assert_eq!(WinRTType::I32.signature_string(), "i4");
        assert_eq!(WinRTType::HString.signature_string(), "string");
        assert_eq!(WinRTType::Bool.signature_string(), "b1");
        assert_eq!(WinRTType::F64.signature_string(), "f8");
        assert_eq!(WinRTType::Guid.signature_string(), "g16");

        let sig = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IASYNC_OPERATION, arity: 1 }), vec![WinRTType::HString]);
        assert_eq!(
            sig.signature_string(),
            "pinterface({9fc2b0bb-e446-44e2-aa61-9cab8f636af2};string)"
        );
    }

    #[test]
    fn test_iid() {
        let iid = GUID::from_u128(0x12345678_1234_1234_1234_123456789abc);
        assert_eq!(WinRTType::Interface(iid).iid(), Some(iid));
        assert_eq!(WinRTType::Delegate(iid).iid(), Some(iid));
        assert_eq!(WinRTType::RuntimeClass("Test".into(), iid).iid(), Some(iid));
        assert_eq!(WinRTType::I32.iid(), None);

        // Parameterized computes IID on demand
        let parameterized_type = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IASYNC_OPERATION, arity: 1 }), vec![WinRTType::HString]);
        let computed = parameterized_type.iid().unwrap();
        let expected = windows_future::IAsyncOperation::<windows_core::HSTRING>::IID;
        assert_eq!(computed, expected);
    }

    #[test]
    fn test_guid_braced_format() {
        let guid = GUID::from_u128(0x9fc2b0bb_e446_44e2_aa61_9cab8f636af2);
        assert_eq!(
            format_guid_braced(&guid),
            "{9fc2b0bb-e446-44e2-aa61-9cab8f636af2}"
        );
    }

    #[test]
    fn test_iid_async_operation_hstring() {
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IASYNC_OPERATION, arity: 1 }), vec![WinRTType::HString]);
        let expected = windows_future::IAsyncOperation::<windows_core::HSTRING>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }

    #[test]
    fn test_iid_async_operation_bool() {
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IASYNC_OPERATION, arity: 1 }), vec![WinRTType::Bool]);
        let expected = windows_future::IAsyncOperation::<bool>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }

    #[test]
    fn test_iid_async_operation_i32() {
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IASYNC_OPERATION, arity: 1 }), vec![WinRTType::I32]);
        let expected = windows_future::IAsyncOperation::<i32>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }

    #[test]
    fn test_iid_vector_hstring() {
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IVECTOR, arity: 1 }), vec![WinRTType::HString]);
        let expected = windows_collections::IVector::<windows_core::HSTRING>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }

    #[test]
    fn test_iid_vector_view_hstring() {
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IVECTOR_VIEW, arity: 1 }), vec![WinRTType::HString]);
        let expected = windows_collections::IVectorView::<windows_core::HSTRING>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }

    #[test]
    fn test_iid_iterable_hstring() {
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IITERABLE, arity: 1 }), vec![WinRTType::HString]);
        let expected = windows_collections::IIterable::<windows_core::HSTRING>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }

    #[test]
    fn test_iid_reference_i32() {
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IREFERENCE, arity: 1 }), vec![WinRTType::I32]);
        let expected = windows::Foundation::IReference::<i32>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }

    #[test]
    fn test_iid_nested_parameterized() {
        // IVector<IVector<HSTRING>>
        let inner = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IVECTOR, arity: 1 }), vec![WinRTType::HString]);
        let outer = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IVECTOR, arity: 1 }), vec![inner]);
        let computed = outer.iid().unwrap();
        assert_eq!((computed.data3 >> 12) & 0xF, 5, "UUID version should be 5");
        assert_eq!((computed.data4[0] >> 6) & 0x3, 2, "UUID variant should be RFC4122");
    }

    #[test]
    fn test_iid_runtime_class_as_type_arg() {
        // IAsyncOperation<StorageFile>
        let storage_file = WinRTType::RuntimeClass(
            "Windows.Storage.StorageFile".into(),
            GUID::from_u128(0xFA3F6186_4214_428C_A64C_14C9AC7315EA),
        );
        let ty = WinRTType::Parameterized(Box::new(WinRTType::Generic { piid: IASYNC_OPERATION, arity: 1 }), vec![storage_file]);
        let expected = windows_future::IAsyncOperation::<windows::Storage::StorageFile>::IID;
        assert_eq!(ty.iid().unwrap(), expected);
    }
}
