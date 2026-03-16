//! Dynamic WinRT delegate (callback) implementation.
//!
//! A delegate is a COM object with IUnknown + a single `Invoke` method.
//! `DynamicDelegate` creates such objects at runtime, marshalling ABI
//! parameters to `WinRTValue` and forwarding to a user-supplied callback.

use core::ffi::c_void;
use windows_core::{GUID, HRESULT, IUnknown, Interface};

use crate::metadata_table::TypeHandle;
use crate::value::WinRTValue;

// ======================================================================
// DynamicDelegate — general-purpose WinRT delegate COM object
// ======================================================================

/// Callback type: receives marshalled Invoke arguments, returns HRESULT.
pub type DelegateCallback = Box<dyn Fn(&[WinRTValue]) -> HRESULT + Send + Sync>;

/// Vtable for a delegate with 2 pointer-sized ABI params (covers ~95% of delegates).
#[repr(C)]
struct Delegate2Vtbl {
    base: windows_core::IUnknown_Vtbl,
    invoke: unsafe extern "system" fn(*mut c_void, *mut c_void, *mut c_void) -> HRESULT,
}

/// A dynamically-constructed WinRT delegate COM object.
///
/// Supports delegates with up to 2 ABI parameters (pointer-sized).
/// This covers TypedEventHandler<T,U>, AsyncOperationCompletedHandler<T>,
/// EventHandler<T>, and most other standard delegates.
#[repr(C)]
struct DynamicDelegate {
    vtable: *const Delegate2Vtbl,
    ref_count: windows_core::imp::RefCount,
    delegate_iid: GUID,
    param_types: Vec<TypeHandle>,
    callback: DelegateCallback,
}

// Safety: DynamicDelegate is ref-counted and the callback is Send+Sync.
unsafe impl Send for DynamicDelegate {}
unsafe impl Sync for DynamicDelegate {}

impl DynamicDelegate {
    const VTBL: Delegate2Vtbl = Delegate2Vtbl {
        base: windows_core::IUnknown_Vtbl {
            QueryInterface: Self::qi,
            AddRef: Self::add_ref,
            Release: Self::release,
        },
        invoke: Self::invoke_2,
    };

    /// Create a new dynamic delegate as an IUnknown COM pointer.
    ///
    /// - `delegate_iid`: the IID of the delegate interface (for QueryInterface)
    /// - `param_types`: types of the Invoke method's parameters (excluding `this`)
    /// - `callback`: function called when WinRT invokes the delegate
    pub fn create(
        delegate_iid: GUID,
        param_types: Vec<TypeHandle>,
        callback: DelegateCallback,
    ) -> IUnknown {
        assert!(
            param_types.len() <= 2,
            "DynamicDelegate currently supports up to 2 parameters, got {}",
            param_types.len()
        );
        let delegate = Box::new(Self {
            vtable: &Self::VTBL,
            ref_count: windows_core::imp::RefCount::new(1),
            delegate_iid,
            param_types,
            callback,
        });
        unsafe { IUnknown::from_raw(Box::into_raw(delegate) as *mut c_void) }
    }

    // ------------------------------------------------------------------
    // IUnknown
    // ------------------------------------------------------------------

    unsafe extern "system" fn qi(
        this: *mut c_void,
        iid: *const GUID,
        ppv: *mut *mut c_void,
    ) -> HRESULT {
        if iid.is_null() || ppv.is_null() {
            return HRESULT(-2147467261); // E_INVALIDARG
        }
        let iid = unsafe { &*iid };
        let delegate = unsafe { &*(this as *const Self) };
        if *iid == IUnknown::IID
            || *iid == windows_core::imp::IAgileObject::IID
            || *iid == delegate.delegate_iid
        {
            unsafe { *ppv = this };
            unsafe { Self::add_ref(this) };
            HRESULT(0)
        } else if *iid == windows_core::imp::IMarshal::IID {
            unsafe {
                delegate.ref_count.add_ref();
                windows_core::imp::marshaler(
                    core::mem::transmute(this),
                    ppv,
                )
            }
        } else {
            unsafe { *ppv = std::ptr::null_mut() };
            HRESULT(-2147467262) // E_NOINTERFACE
        }
    }

    unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
        let delegate = unsafe { &*(this as *const Self) };
        delegate.ref_count.add_ref()
    }

    unsafe extern "system" fn release(this: *mut c_void) -> u32 {
        let delegate = unsafe { &*(this as *const Self) };
        let remaining = delegate.ref_count.release();
        if remaining == 0 {
            unsafe { drop(Box::from_raw(this as *mut Self)) };
        }
        remaining
    }

    // ------------------------------------------------------------------
    // Invoke trampoline (2 pointer-sized ABI params)
    // ------------------------------------------------------------------

    unsafe extern "system" fn invoke_2(
        this: *mut c_void,
        arg0: *mut c_void,
        arg1: *mut c_void,
    ) -> HRESULT {
        let delegate = unsafe { &*(this as *const Self) };
        let raw_args = [arg0, arg1];
        let mut values = Vec::with_capacity(delegate.param_types.len());

        for (i, pt) in delegate.param_types.iter().enumerate() {
            if i < raw_args.len() {
                values.push(marshal_abi_ptr(raw_args[i], pt));
            }
        }

        (delegate.callback)(&values)
    }
}

/// Convert a raw ABI pointer-sized argument to WinRTValue, based on type.
fn marshal_abi_ptr(raw: *mut c_void, typ: &TypeHandle) -> WinRTValue {
    use crate::metadata_table::TypeKind;
    match typ.kind() {
        // Pointer-sized types: wrap as Object (AddRef via from_raw_borrowed + clone)
        TypeKind::Object | TypeKind::Interface(_) | TypeKind::RuntimeClass(_)
        | TypeKind::Delegate(_) | TypeKind::Parameterized(_) => {
            if raw.is_null() {
                WinRTValue::Null
            } else {
                let obj = unsafe { IUnknown::from_raw_borrowed(&raw) }.unwrap();
                WinRTValue::Object(obj.clone())
            }
        }
        // HString: transmute the raw HSTRING handle
        TypeKind::HString => {
            if raw.is_null() {
                WinRTValue::HString(windows_core::HSTRING::new())
            } else {
                let hstr: &windows_core::HSTRING = unsafe {
                    &*(&raw as *const *mut c_void as *const windows_core::HSTRING)
                };
                WinRTValue::HString(hstr.clone())
            }
        }
        // Small integer types packed into pointer-sized arg
        TypeKind::Bool => WinRTValue::Bool((raw as usize) != 0),
        TypeKind::I32 => WinRTValue::I32(raw as i32),
        TypeKind::Enum(_) => WinRTValue::Enum { value: raw as i32, type_handle: typ.clone() },
        TypeKind::U32 => WinRTValue::U32(raw as u32),
        TypeKind::I64 => WinRTValue::I64(raw as i64),
        TypeKind::U64 => WinRTValue::U64(raw as u64),
        _ => {
            // Fallback: treat as raw i64 (covers most ABI-compatible cases)
            WinRTValue::I64(raw as i64)
        }
    }
}

// ======================================================================
// Public API
// ======================================================================

/// Create a dynamic WinRT delegate COM object.
///
/// # Arguments
/// - `delegate_iid`: the delegate interface IID
/// - `param_types`: Invoke parameter types (max 2)
/// - `callback`: called when WinRT invokes the delegate
///
/// # Returns
/// An `IUnknown` smart pointer to the delegate COM object.
/// Pass this to WinRT methods that accept the delegate (e.g. event subscriptions).
pub fn create_delegate(
    delegate_iid: GUID,
    param_types: Vec<TypeHandle>,
    callback: DelegateCallback,
) -> IUnknown {
    DynamicDelegate::create(delegate_iid, param_types, callback)
}

/// Convenience: create a delegate and wrap as WinRTValue::Object.
pub fn create_delegate_value(
    delegate_iid: GUID,
    param_types: Vec<TypeHandle>,
    callback: DelegateCallback,
) -> WinRTValue {
    WinRTValue::Object(create_delegate(delegate_iid, param_types, callback))
}
