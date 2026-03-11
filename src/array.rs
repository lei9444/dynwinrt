use core::ffi::c_void;
use windows_core::{IUnknown, Interface};

use crate::metadata_table::{TypeHandle, TypeKind};
use crate::value::WinRTValue;


/// How the backing buffer was allocated — determines deallocation strategy.
enum ArrayBuffer {
    /// Owns the buffer AND the element references.
    /// Used for ReceiveArray (callee-allocated) and FillArray (caller-allocated via CoTaskMemAlloc).
    /// Drop releases non-blittable elements, then CoTaskMemFree.
    CoTaskMem(*mut c_void),
    /// Borrowed view for PassArray — buffer owns the bytes but NOT the element references.
    /// Drop frees the Vec but does NOT Release/DeleteString elements.
    /// The original WinRTValues retain ownership of the references.
    Borrowed(Vec<u8>),
}

impl ArrayBuffer {
    /// Whether this buffer owns the element references (needs per-element cleanup on drop).
    fn owns_elements(&self) -> bool {
        matches!(self, ArrayBuffer::CoTaskMem(_))
    }
}

impl std::fmt::Debug for ArrayBuffer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ArrayBuffer::CoTaskMem(p) => write!(f, "CoTaskMem({:p})", p),
            ArrayBuffer::Borrowed(v) => write!(f, "Borrowed({} bytes)", v.len()),
        }
    }
}

/// Holds a dynamically-typed WinRT array backed by a raw byte buffer.
///
/// For blittable types, elements can be read as zero-copy slices via `as_slice::<T>()`.
/// For non-blittable types (HString, Object), individual element access returns cloned values.
/// Drop handles releasing non-blittable elements (Release / WindowsDeleteString) before
/// freeing the buffer itself.
pub struct ArrayData {
    pub element_type: TypeHandle,
    len: usize,
    buffer: ArrayBuffer,
}

impl std::fmt::Debug for ArrayData {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ArrayData")
            .field("element_type", &self.element_type)
            .field("len", &self.len)
            .field("buffer", &self.buffer)
            .finish()
    }
}

impl ArrayData {
    pub fn empty(element_type: TypeHandle) -> Self {
        ArrayData {
            element_type,
            len: 0,
            buffer: ArrayBuffer::CoTaskMem(std::ptr::null_mut()),
        }
    }

    /// Create an ArrayData by serializing WinRTValues into a contiguous buffer.
    /// Used for PassArray in-params.
    ///
    /// The buffer borrows element references (no AddRef/Duplicate).
    /// The caller's WinRTValues retain ownership — they must outlive this ArrayData.
    pub fn from_values(element_type: TypeHandle, values: &[WinRTValue]) -> Self {
        let buffer = serialize_to_buffer(&element_type, values);
        ArrayData {
            len: values.len(),
            element_type,
            buffer: ArrayBuffer::Borrowed(buffer),
        }
    }

    /// Wrap a CoTaskMem-allocated buffer (ReceiveArray or FillArray pattern).
    /// ArrayData takes ownership and will CoTaskMemFree on drop.
    pub(crate) fn from_cotaskmem(
        element_type: TypeHandle,
        data_ptr: *mut c_void,
        len: usize,
    ) -> Self {
        ArrayData {
            element_type,
            len,
            buffer: ArrayBuffer::CoTaskMem(data_ptr),
        }
    }

    pub fn len(&self) -> usize {
        self.len
    }

    fn data_ptr(&self) -> *const u8 {
        match &self.buffer {
            ArrayBuffer::Borrowed(v) => v.as_ptr(),
            ArrayBuffer::CoTaskMem(p) => *p as *const u8,
        }
    }

    // ------------------------------------------------------------------
    // Blittable element access — zero-copy slice
    // ------------------------------------------------------------------

    /// Return the raw buffer as a typed slice. Only valid for blittable types
    /// where `size_of::<T>() == element_type.element_size()`.
    ///
    /// # Safety
    /// Caller must ensure T matches the actual element layout.
    pub unsafe fn as_typed_slice<T: Copy>(&self) -> &[T] {
        assert_eq!(
            std::mem::size_of::<T>(),
            self.element_type.element_size(),
            "as_typed_slice<T> size mismatch"
        );
        if self.len == 0 {
            return &[];
        }
        unsafe { std::slice::from_raw_parts(self.data_ptr() as *const T, self.len) }
    }

    // ------------------------------------------------------------------
    // Per-element access (works for all types)
    // ------------------------------------------------------------------

    /// Read element at `index` as a WinRTValue.
    /// For non-blittable types this clones (AddRef / DuplicateString).
    pub fn get(&self, index: usize) -> WinRTValue {
        assert!(index < self.len, "ArrayData::get index {} out of bounds (len {})", index, self.len);
        let elem_size = self.element_type.element_size();
        let base = self.data_ptr();
        unsafe {
            match self.element_type.kind() {
                TypeKind::Bool => {
                    WinRTValue::Bool(*base.add(index * elem_size) != 0)
                }
                TypeKind::I8 => {
                    WinRTValue::I8(*(base.add(index * elem_size) as *const i8))
                }
                TypeKind::U8 => {
                    WinRTValue::U8(*base.add(index * elem_size))
                }
                TypeKind::I16 => {
                    WinRTValue::I16(*(base.add(index * elem_size) as *const i16))
                }
                TypeKind::U16 | TypeKind::Char16 => {
                    WinRTValue::U16(*(base.add(index * elem_size) as *const u16))
                }
                TypeKind::I32 => {
                    WinRTValue::I32(*(base.add(index * elem_size) as *const i32))
                }
                TypeKind::U32 => {
                    WinRTValue::U32(*(base.add(index * elem_size) as *const u32))
                }
                TypeKind::I64 => {
                    WinRTValue::I64(*(base.add(index * elem_size) as *const i64))
                }
                TypeKind::U64 => {
                    WinRTValue::U64(*(base.add(index * elem_size) as *const u64))
                }
                TypeKind::F32 => {
                    WinRTValue::F32(*(base.add(index * elem_size) as *const f32))
                }
                TypeKind::F64 => {
                    WinRTValue::F64(*(base.add(index * elem_size) as *const f64))
                }
                TypeKind::Guid => {
                    let guid = *(base.add(index * 16) as *const windows_core::GUID);
                    WinRTValue::Guid(guid)
                }
                TypeKind::HString => {
                    let raw = *(base.add(index * elem_size) as *const *mut c_void);
                    // Duplicate: read the handle and clone it (bumps refcount)
                    let hstr: &windows_core::HSTRING = &*((&raw) as *const *mut c_void as *const windows_core::HSTRING);
                    WinRTValue::HString(hstr.clone())
                }
                kind if kind.is_com_pointer() => {
                    let raw = *(base.add(index * elem_size) as *const *mut c_void);
                    if raw.is_null() {
                        WinRTValue::Object(IUnknown::from_raw(std::ptr::null_mut()))
                    } else {
                        // from_raw takes ownership, but we want a clone — so AddRef first
                        let obj = IUnknown::from_raw_borrowed(&raw).unwrap();
                        WinRTValue::Object(obj.clone())
                    }
                }
                TypeKind::Struct(_) => {
                    let sz = self.element_type.size_of();
                    let mut vd = self.element_type.default_value();
                    std::ptr::copy_nonoverlapping(
                        base.add(index * sz),
                        vd.as_mut_ptr(),
                        sz,
                    );
                    WinRTValue::Struct(vd)
                }
                other => panic!("ArrayData::get unsupported element type: {:?}", other),
            }
        }
    }

    // ------------------------------------------------------------------
    // Convenience typed getters
    // ------------------------------------------------------------------

    pub fn get_i32(&self, index: usize) -> i32 {
        assert!(index < self.len);
        unsafe { *(self.data_ptr().add(index * 4) as *const i32) }
    }

    /// Return the raw buffer pointer and length for passing to WinRT ABI (PassArray).
    pub(crate) fn as_raw_parts(&self) -> (*const u8, usize) {
        (self.data_ptr(), self.len)
    }
}

// ======================================================================
// Drop — release non-blittable elements, then free the buffer
// ======================================================================

impl Drop for ArrayData {
    fn drop(&mut self) {
        // Swap out to prevent double-free if element destructor panics
        let mut buffer = ArrayBuffer::CoTaskMem(std::ptr::null_mut());
        let mut len = 0usize;
        std::mem::swap(&mut buffer, &mut self.buffer);
        std::mem::swap(&mut len, &mut self.len);

        if len > 0 && buffer.owns_elements() {
            let base = match &buffer {
                ArrayBuffer::CoTaskMem(p) => *p as *const u8,
                ArrayBuffer::Borrowed(v) => v.as_ptr(),
            };
            let elem_size = self.element_type.element_size();
            let kind = self.element_type.kind();

            // Release non-blittable elements
            match kind {
                TypeKind::HString => {
                    for i in 0..len {
                        unsafe {
                            let raw = *(base.add(i * elem_size) as *const *mut c_void);
                            if !raw.is_null() {
                                // WindowsDeleteString by transmuting to HSTRING and dropping
                                let _hstr: windows_core::HSTRING = std::mem::transmute(raw);
                                // _hstr drops here, calling WindowsDeleteString
                            }
                        }
                    }
                }
                kind if kind.is_com_pointer() => {
                    for i in 0..len {
                        unsafe {
                            let raw = *(base.add(i * elem_size) as *const *mut c_void);
                            if !raw.is_null() {
                                // from_raw takes ownership → drop calls Release
                                let _obj = IUnknown::from_raw(raw);
                            }
                        }
                    }
                }
                _ => {
                    // Blittable types (including Struct): no per-element cleanup needed
                }
            }
        }

        // Free the buffer
        match buffer {
            ArrayBuffer::Borrowed(_) => {
                // Vec drops naturally
            }
            ArrayBuffer::CoTaskMem(ptr) => {
                if !ptr.is_null() {
                    unsafe {
                        windows::Win32::System::Com::CoTaskMemFree(Some(ptr));
                    }
                }
            }
        }
    }
}

// ======================================================================
// Clone — deep copy with AddRef/DuplicateString for non-blittable
// ======================================================================

impl Clone for ArrayData {
    fn clone(&self) -> Self {
        if self.len == 0 {
            return ArrayData::empty(self.element_type.clone());
        }

        let elem_size = self.element_type.element_size();
        let total_bytes = self.len * elem_size;
        let base = self.data_ptr();

        // Borrowed buffer: raw memcpy, stays Borrowed (no ownership to duplicate).
        if !self.buffer.owns_elements() {
            let mut buf = vec![0u8; total_bytes];
            unsafe { std::ptr::copy_nonoverlapping(base, buf.as_mut_ptr(), total_bytes) };
            return ArrayData {
                element_type: self.element_type.clone(),
                len: self.len,
                buffer: ArrayBuffer::Borrowed(buf),
            };
        }

        // CoTaskMem: must AddRef/Duplicate non-blittable elements.
        // Allocate new CoTaskMem buffer for the clone.
        let new_ptr = unsafe {
            windows::Win32::System::Com::CoTaskMemAlloc(total_bytes)
        };
        assert!(!new_ptr.is_null(), "CoTaskMemAlloc failed in ArrayData::clone");
        let new_buf = new_ptr as *mut u8;

        let kind = self.element_type.kind();
        match kind {
            TypeKind::HString => {
                unsafe { std::ptr::write_bytes(new_buf, 0, total_bytes) };
                for i in 0..self.len {
                    unsafe {
                        let raw = *(base.add(i * elem_size) as *const *mut c_void);
                        if !raw.is_null() {
                            let hstr: &windows_core::HSTRING = &*((&raw) as *const *mut c_void as *const windows_core::HSTRING);
                            let cloned: *mut c_void = std::mem::transmute(hstr.clone());
                            (new_buf.add(i * elem_size) as *mut *mut c_void).write(cloned);
                        }
                    }
                }
            }
            kind if kind.is_com_pointer() => {
                unsafe { std::ptr::write_bytes(new_buf, 0, total_bytes) };
                for i in 0..self.len {
                    unsafe {
                        let raw = *(base.add(i * elem_size) as *const *mut c_void);
                        if !raw.is_null() {
                            let obj = IUnknown::from_raw_borrowed(&raw).unwrap().clone();
                            (new_buf.add(i * elem_size) as *mut *mut c_void).write(obj.into_raw());
                        }
                    }
                }
            }
            _ => {
                // Blittable (primitives, Guid, Struct): plain memcpy
                unsafe { std::ptr::copy_nonoverlapping(base, new_buf, total_bytes) };
            }
        }

        ArrayData {
            element_type: self.element_type.clone(),
            len: self.len,
            buffer: ArrayBuffer::CoTaskMem(new_ptr),
        }
    }
}

// ======================================================================
// Serialization — used for PassArray (in-params from WinRTValue)
// ======================================================================

/// Serialize WinRTValue elements into a contiguous byte buffer for PassArray ABI.
/// This is also used by `ArrayData::from_values`.
fn serialize_to_buffer(element_type: &TypeHandle, values: &[WinRTValue]) -> Vec<u8> {
    let elem_size = element_type.element_size();
    let mut buffer = Vec::with_capacity(values.len() * elem_size);
    for elem in values {
        match elem {
            WinRTValue::Bool(v) => buffer.push(*v as u8),
            WinRTValue::I8(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::U8(v) => buffer.push(*v),
            WinRTValue::I16(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::U16(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::I32(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::U32(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::I64(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::U64(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::F32(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::F64(v) => buffer.extend_from_slice(&v.to_ne_bytes()),
            WinRTValue::Object(obj) => {
                buffer.extend_from_slice(&(obj.as_raw() as usize).to_ne_bytes());
            }
            WinRTValue::HString(s) => {
                let raw: usize = unsafe { std::mem::transmute_copy(s) };
                buffer.extend_from_slice(&raw.to_ne_bytes());
            }
            WinRTValue::Guid(g) => {
                let bytes: &[u8; 16] = unsafe { &*(g as *const windows_core::GUID as *const [u8; 16]) };
                buffer.extend_from_slice(bytes);
            }
            WinRTValue::Struct(vd) => {
                let size = vd.type_handle().size_of();
                let src = unsafe { std::slice::from_raw_parts(vd.as_ptr(), size) };
                buffer.extend_from_slice(src);
            }
            _ => panic!("Unsupported array element type for serialization: {:?}", elem),
        }
    }
    buffer
}

/// Serialize array elements for the PassArray ABI pattern.
/// Returns a byte buffer that must outlive the FFI call.
pub(crate) fn serialize_array_elements(data: &ArrayData) -> (*const u8, usize) {
    data.as_raw_parts()
}

#[cfg(test)]
mod tests {
    use std::alloc::Layout;

    use libffi::low::CodePtr;
    use windows::Win32::System::WinRT::{
        IActivationFactory, RO_INIT_MULTITHREADED, RoGetActivationFactory, RoInitialize,
    };
    use windows_core::HRESULT;

    use crate::call::get_vtable_function_ptr;

    #[test]
    fn cryptographic_buffer_test() -> windows::core::Result<()> {
        use windows::Security::Cryptography::CryptographicBuffer;
        let value = vec![1u8, 2, 3, 4, 5, 6, 7, 8, 9, 10];
        let buffer = CryptographicBuffer::CreateFromByteArray(&value)?;
        let base64 = CryptographicBuffer::EncodeToBase64String(&buffer)?;
        println!("Generated base64 string: {}", base64);
        Ok(())
    }

    #[tokio::test]
    async fn geolocation_value_type_test() -> windows::core::Result<()> {
        use windows::Devices::Geolocation::{BasicGeoposition, Geopoint, Geoposition};
        let position = BasicGeoposition {
            Latitude: 47.643,
            Longitude: -122.131,
            Altitude: 0.0,
        };
        let geopoint = Geopoint::Create(position)?;
        println!(
            "Geopoint created at lat: {}, lon: {}",
            geopoint.Position()?.Latitude,
            geopoint.Position()?.Longitude
        );
        let locator = windows::Devices::Geolocation::Geolocator::new()?;
        let geoposition: Geoposition = locator.GetGeopositionAsync()?.await?;
        println!(
            "Current location: lat: {}, lon: {}",
            geoposition.Coordinate()?.Point()?.Position()?.Latitude,
            geoposition.Coordinate()?.Point()?.Position()?.Longitude
        );
        Ok(())
    }

    #[test]
    fn geolocation_value_type_dynamic() -> windows::core::Result<()> {
        use windows::Devices::Geolocation::{BasicGeoposition, Geopoint};
        use windows::core::h;
        use windows::core::{IInspectable, Interface};

        unsafe {
            RoInitialize(RO_INIT_MULTITHREADED);
        }
        let position = BasicGeoposition {
            Latitude: 47.643,
            Longitude: -122.131,
            Altitude: 0.0,
        };
        let afactory = unsafe {
            RoGetActivationFactory::<IActivationFactory>(h!("Windows.Devices.Geolocation.Geopoint"))
        }?;
        let GeopointFactory = afactory.cast::<windows::Devices::Geolocation::IGeopointFactory>()?;
        let createFptr = get_vtable_function_ptr(GeopointFactory.as_raw(), 6);
        let create = unsafe {
            std::mem::transmute::<
                _,
                unsafe extern "system" fn(
                    *mut std::ffi::c_void,
                    BasicGeoposition,
                    *mut *mut std::ffi::c_void,
                ) -> windows::core::HRESULT,
            >(createFptr)
        };
        let mut out = std::ptr::null_mut();
        let hr = unsafe { create(GeopointFactory.as_raw(), position, &mut out) };
        hr.ok()?;
        let geopoint = unsafe { Geopoint::from_raw(out) };
        let inspectable: IInspectable = geopoint.cast()?;
        let dynamic_geopoint: Geopoint = inspectable.cast()?;
        println!(
            "Dynamic Geopoint created at lat: {}, lon: {}",
            dynamic_geopoint.Position()?.Latitude,
            dynamic_geopoint.Position()?.Longitude
        );
        Ok(())
    }

    #[test]
    fn geolocation_value_type_dynamic_libffi() -> windows::core::Result<()> {
        use windows::Devices::Geolocation::{BasicGeoposition, Geopoint};
        use windows::core::h;
        use windows::core::{IInspectable, Interface};

        unsafe {
            RoInitialize(RO_INIT_MULTITHREADED);
        }
        let position = BasicGeoposition {
            Latitude: 47.643,
            Longitude: -122.131,
            Altitude: 0.0,
        };
        let afactory = unsafe {
            RoGetActivationFactory::<IActivationFactory>(h!("Windows.Devices.Geolocation.Geopoint"))
        }?;

        let BasicGeoPositionStruct = libffi::middle::Type::structure(vec![
            libffi::middle::Type::f64(),
            libffi::middle::Type::f64(),
            libffi::middle::Type::f64(),
        ]);
        let f1 = Layout::new::<f64>();
        let (f2, offset2) = f1.extend(Layout::new::<f64>()).unwrap();
        let (f3, offset3) = f2.extend(Layout::new::<f64>()).unwrap();
        let sl = f3.pad_to_align();

        let sptr = unsafe { std::alloc::alloc(sl) };
        let pf1 = unsafe { sptr } as *mut f64;
        let pf2 = unsafe { sptr.add(offset2) } as *mut f64;
        let pf3 = unsafe { sptr.add(offset3) } as *mut f64;
        unsafe {
            *pf1 = 11.0;
            *pf2 = 22.0;
            *pf3 = 33.0;
        }

        let create = libffi::middle::Cif::new(
            vec![
                libffi::middle::Type::pointer(),
                BasicGeoPositionStruct,
                libffi::middle::Type::pointer(),
            ]
            .into_iter(),
            libffi::middle::Type::i32(),
        );
        let GeopointFactory = afactory.cast::<windows::Devices::Geolocation::IGeopointFactory>()?;
        let createFptr = get_vtable_function_ptr(GeopointFactory.as_raw(), 6);
        let mut out = std::ptr::null_mut();
        let pOut = &mut out as *mut *mut std::ffi::c_void;
        let hr = unsafe {
            libffi::low::call::<HRESULT>(
                create.as_raw_ptr(),
                CodePtr(createFptr),
                vec![
                    &GeopointFactory.as_raw() as *const _ as *mut std::ffi::c_void,
                    sptr as *const _ as *mut std::ffi::c_void,
                    &pOut as *const _ as *mut std::ffi::c_void,
                ]
                .as_mut_ptr(),
            )
        };
        hr.ok()?;
        let geopoint = unsafe { Geopoint::from_raw(out) };
        let inspectable: IInspectable = geopoint.cast()?;
        let dynamic_geopoint: Geopoint = inspectable.cast()?;
        println!(
            "Dynamic Geopoint created at lat: {}, lon: {}",
            dynamic_geopoint.Position()?.Latitude,
            dynamic_geopoint.Position()?.Longitude
        );
        Ok(())
    }

    #[tokio::test]
    async fn enumerate_device_test() -> windows::core::Result<()> {
        use windows::Devices::Enumeration::DeviceInformation;
        let devices = DeviceInformation::FindAllAsync()?.await?;
        let mut items = windows::core::Array::<DeviceInformation>::with_len(30);
        let count = devices.GetMany(10, &mut items)?;
        println!("Found {} devices", count);
        for device in items[..count as usize].iter() {
            println!(
                "Device: {} - {}",
                device.as_ref().unwrap().Name()?,
                device.as_ref().unwrap().Id()?
            );
        }
        Ok(())
    }
}
