use std::future::{Future, IntoFuture};
use std::pin::Pin;
use std::sync::{Arc, Mutex};
use std::task::{Context, Poll, Waker};

use windows::core::Interface;
use windows_core::{GUID, HRESULT, IUnknown};
use windows_future::{AsyncActionCompletedHandler, AsyncStatus};

use crate::result::{Error, Result};
use crate::types::WinRTType;
use crate::value::WinRTValue;

// ---------------------------------------------------------------------------
// DynCompletedHandler — a minimal COM object for WinRT completion callbacks
// Used for generic async types (IAsyncOperation<T>, etc.) where the handler
// IID is parameterized and we can't use windows-future's typed handlers.
// ---------------------------------------------------------------------------

#[repr(C)]
struct DynCompletedHandlerVtbl {
    base: windows_core::IUnknown_Vtbl,
    invoke: unsafe extern "system" fn(
        this: *mut std::ffi::c_void,
        sender: *mut std::ffi::c_void,
        status: AsyncStatus,
    ) -> HRESULT,
}

#[repr(C)]
struct DynCompletedHandler {
    vtable: *const DynCompletedHandlerVtbl,
    ref_count: windows_core::imp::RefCount,
    handler_iid: GUID,
    waker: Arc<Mutex<Waker>>,
}

impl DynCompletedHandler {
    const VTBL: DynCompletedHandlerVtbl = DynCompletedHandlerVtbl {
        base: windows_core::IUnknown_Vtbl {
            QueryInterface: Self::qi,
            AddRef: Self::add_ref,
            Release: Self::release,
        },
        invoke: Self::invoke,
    };

    fn create(waker: Arc<Mutex<Waker>>, handler_iid: GUID) -> IUnknown {
        let handler = Box::new(Self {
            vtable: &Self::VTBL,
            ref_count: windows_core::imp::RefCount::new(1),
            handler_iid,
            waker,
        });
        unsafe { IUnknown::from_raw(Box::into_raw(handler) as *mut std::ffi::c_void) }
    }

    unsafe extern "system" fn qi(
        this: *mut std::ffi::c_void,
        iid: *const GUID,
        ppv: *mut *mut std::ffi::c_void,
    ) -> HRESULT {
        if iid.is_null() || ppv.is_null() {
            return HRESULT(-2147467261); // E_INVALIDARG
        }
        let iid = unsafe { &*iid };
        let handler = unsafe { &*(this as *const Self) };
        if *iid == IUnknown::IID
            || *iid == windows_core::imp::IAgileObject::IID
            || *iid == handler.handler_iid
        {
            unsafe { *ppv = this };
            unsafe { Self::add_ref(this) };
            HRESULT(0) // S_OK
        } else if *iid == windows_core::imp::IMarshal::IID {
            unsafe {
                handler.ref_count.add_ref();
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

    unsafe extern "system" fn add_ref(this: *mut std::ffi::c_void) -> u32 {
        let handler = unsafe { &*(this as *const Self) };
        handler.ref_count.add_ref()
    }

    unsafe extern "system" fn release(this: *mut std::ffi::c_void) -> u32 {
        let handler = unsafe { &*(this as *const Self) };
        let remaining = handler.ref_count.release();
        if remaining == 0 {
            unsafe { drop(Box::from_raw(this as *mut Self)) };
        }
        remaining
    }

    unsafe extern "system" fn invoke(
        this: *mut std::ffi::c_void,
        _sender: *mut std::ffi::c_void,
        _status: AsyncStatus,
    ) -> HRESULT {
        let handler = unsafe { &*(this as *const Self) };
        if let Ok(waker) = handler.waker.lock() {
            waker.wake_by_ref();
        }
        HRESULT(0) // S_OK
    }
}

// ---------------------------------------------------------------------------
// WinRTAsyncFuture — event-driven Future for dynamic WinRT async operations
// ---------------------------------------------------------------------------

use crate::value::AsyncInfo;

pub struct WinRTAsyncFuture {
    async_info: AsyncInfo,
    waker: Option<Arc<Mutex<Waker>>>,
}

impl WinRTAsyncFuture {
    fn from_value(value: WinRTValue) -> Self {
        match value {
            WinRTValue::Async(a) => Self { async_info: a, waker: None },
            _ => panic!("WinRTAsyncFuture::from_value called with non-async WinRTValue"),
        }
    }

    /// QI from IAsyncInfo to the concrete async interface.
    fn query_concrete(&self) -> Result<IUnknown> {
        let iid = self.async_info.iid();
        let mut ptr = std::ptr::null_mut();
        unsafe { self.async_info.info.query(&iid, &mut ptr) }
            .ok()
            .map_err(Error::WindowsError)?;
        Ok(unsafe { IUnknown::from_raw(ptr) })
    }

    /// Vtable indices for SetCompleted and GetResults.
    fn vtable_indices(&self) -> (usize, usize) {
        match &self.async_info.async_type {
            WinRTType::IAsyncAction | WinRTType::IAsyncOperation(_) => (6, 8),
            WinRTType::IAsyncActionWithProgress(_) | WinRTType::IAsyncOperationWithProgress(_, _) => (8, 10),
            _ => panic!("not an async type"),
        }
    }

    /// Call GetResults on the concrete async interface and return the WinRTValue.
    fn get_results(&self) -> Result<WinRTValue> {
        let concrete = self.query_concrete()?;
        let (_, get_results_index) = self.vtable_indices();

        if let Some(rt) = self.async_info.result_type() {
            let mut out = std::ptr::null_mut();
            let hr = crate::call::call_winrt_method_1(
                get_results_index,
                concrete.as_raw(),
                &mut out,
            );
            hr.ok().map_err(Error::WindowsError)?;
            rt.from_out(out)
        } else {
            let mut dummy: *mut std::ffi::c_void = std::ptr::null_mut();
            let hr = crate::call::call_winrt_method_1(
                get_results_index,
                concrete.as_raw(),
                &mut dummy,
            );
            hr.ok().map_err(Error::WindowsError)?;
            Ok(WinRTValue::HResult(HRESULT(0)))
        }
    }

    /// Register SetCompleted using the typed windows-future API (IAsyncAction)
    /// or via dynamic vtable call (generic types).
    fn register_completed(&self, shared_waker: Arc<Mutex<Waker>>) -> Result<()> {
        if self.async_info.iid() == crate::types::IASYNC_ACTION {
            // IAsyncAction — use windows-future's typed handler directly
            let action: windows_future::IAsyncAction = self.async_info.info.cast()
                .map_err(Error::WindowsError)?;
            let handler = AsyncActionCompletedHandler::new(move |_, _| {
                if let Ok(waker) = shared_waker.lock() {
                    waker.wake_by_ref();
                }
                Ok(())
            });
            action.SetCompleted(&handler)
                .map_err(Error::WindowsError)?;
        } else {
            // Generic types — use DynCompletedHandler via vtable
            let handler = DynCompletedHandler::create(shared_waker, self.async_info.handler_iid());
            let concrete = self.query_concrete()?;
            let (set_completed_index, _) = self.vtable_indices();
            let hr = crate::call::call_winrt_method_1(
                set_completed_index,
                concrete.as_raw(),
                handler.as_raw(),
            );
            hr.ok().map_err(Error::WindowsError)?;
        }
        Ok(())
    }
}

impl Future for WinRTAsyncFuture {
    type Output = Result<WinRTValue>;

    fn poll(mut self: Pin<&mut Self>, cx: &mut Context<'_>) -> Poll<Self::Output> {
        // Fast path: already completed before first poll
        match self.async_info.info.Status() {
            Ok(status) if status != AsyncStatus::Started => {
                return Poll::Ready(self.get_results());
            }
            Err(e) => return Poll::Ready(Err(Error::WindowsError(e))),
            _ => {}
        }

        if let Some(shared_waker) = &self.waker {
            // Subsequent poll — update the waker in case executor changed
            if let Ok(mut guard) = shared_waker.lock() {
                guard.clone_from(cx.waker());
            }
            // Re-check status (race: completion may have fired between status check and here)
            match self.async_info.info.Status() {
                Ok(status) if status != AsyncStatus::Started => {
                    return Poll::Ready(self.get_results());
                }
                Err(e) => return Poll::Ready(Err(Error::WindowsError(e))),
                _ => {}
            }
        } else {
            // First poll — register SetCompleted
            let shared_waker = Arc::new(Mutex::new(cx.waker().clone()));
            self.waker = Some(shared_waker.clone());

            if let Err(e) = self.register_completed(shared_waker) {
                return Poll::Ready(Err(e));
            }
        }

        Poll::Pending
    }
}

// ---------------------------------------------------------------------------
// IntoFuture for WinRTValue
// ---------------------------------------------------------------------------

impl IntoFuture for WinRTValue {
    type Output = Result<WinRTValue>;
    type IntoFuture = WinRTAsyncFuture;

    fn into_future(self) -> WinRTAsyncFuture {
        WinRTAsyncFuture::from_value(self)
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use windows::core::Interface;
    use windows::System::Threading::{ThreadPool, WorkItemHandler};
    use windows_future::IAsyncInfo;

    use crate::result::{Error, Result};
    use crate::value::{AsyncInfo, WinRTValue};
    use crate::types::WinRTType;

    #[tokio::test]
    async fn test_async_action() -> Result<()> {
        // ThreadPool.RunAsync returns IAsyncAction (no type parameters)
        let handler = WorkItemHandler::new(|_| Ok(()));
        let op = ThreadPool::RunAsync(&handler)
            .map_err(Error::WindowsError)?;
        let async_info: IAsyncInfo = op.cast()
            .map_err(Error::WindowsError)?;

        let value = WinRTValue::Async(AsyncInfo {
            info: async_info,
            async_type: WinRTType::IAsyncAction,
        });
        let _result = value.await?;
        println!("IAsyncAction completed successfully");
        Ok(())
    }
}
