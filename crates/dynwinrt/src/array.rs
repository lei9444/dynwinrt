use core::ffi::c_void;
use windows_core::{IUnknown, Interface};

use crate::metadata_table::{TypeHandle, TypeKind};
use crate::value::WinRTValue;


/// How the array data is stored.
enum ArrayBuffer {
    /// User-built array (for PassArray). Elements are owned WinRTValues.
    /// Serialized to raw bytes only at FFI call time.
    Values(Vec<WinRTValue>),
    /// WinRT-allocated buffer (ReceiveArray / FillArray).
    /// Owns the buffer AND the element references.
    /// Drop releases non-blittable elements, then CoTaskMemFree.
    CoTaskMem { ptr: *mut c_void, len: usize },
}

impl std::fmt::Debug for ArrayBuffer {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ArrayBuffer::Values(v) => write!(f, "Values({} elements)", v.len()),
            ArrayBuffer::CoTaskMem { ptr, len } => write!(f, "CoTaskMem({:p}, {} elements)", ptr, len),
        }
    }
}

/// Holds a dynamically-typed WinRT array.
///
/// Two representations:
/// - `Values`: owned `Vec<WinRTValue>`, used for arrays the caller builds (PassArray).
///   Clone/Drop delegate to WinRTValue which handles refcounting automatically.
/// - `CoTaskMem`: raw byte buffer from WinRT (ReceiveArray/FillArray).
///   Clone/Drop manually handle per-element refcounting on raw bytes.
pub struct ArrayData {
    pub element_type: TypeHandle,
    buffer: ArrayBuffer,
}

impl std::fmt::Debug for ArrayData {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ArrayData")
            .field("element_type", &self.element_type)
            .field("buffer", &self.buffer)
            .finish()
    }
}

impl ArrayData {
    pub fn empty(element_type: TypeHandle) -> Self {
        ArrayData {
            element_type,
            buffer: ArrayBuffer::CoTaskMem { ptr: std::ptr::null_mut(), len: 0 },
        }
    }

    /// Create an ArrayData from owned WinRTValues.
    /// Used for PassArray in-params. The values are cloned and owned by this ArrayData.
    pub fn from_values(element_type: TypeHandle, values: &[WinRTValue]) -> Self {
        ArrayData {
            element_type,
            buffer: ArrayBuffer::Values(values.to_vec()),
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
            buffer: ArrayBuffer::CoTaskMem { ptr: data_ptr, len },
        }
    }

    pub fn len(&self) -> usize {
        match &self.buffer {
            ArrayBuffer::Values(v) => v.len(),
            ArrayBuffer::CoTaskMem { len, .. } => *len,
        }
    }

    // ------------------------------------------------------------------
    // Blittable element access — zero-copy slice (CoTaskMem only)
    // ------------------------------------------------------------------

    /// Return the raw buffer as a typed slice. Only valid for CoTaskMem-backed arrays
    /// with blittable types where `size_of::<T>() == element_type.element_size()`.
    ///
    /// # Safety
    /// Caller must ensure T matches the actual element layout.
    pub unsafe fn as_typed_slice<T: Copy>(&self) -> &[T] {
        match &self.buffer {
            ArrayBuffer::CoTaskMem { ptr, len } => {
                assert_eq!(
                    std::mem::size_of::<T>(),
                    self.element_type.element_size(),
                    "as_typed_slice<T> size mismatch"
                );
                if *len == 0 {
                    return &[];
                }
                unsafe { std::slice::from_raw_parts(*ptr as *const T, *len) }
            }
            ArrayBuffer::Values(_) => {
                panic!("as_typed_slice not available for Values arrays; use get() instead")
            }
        }
    }

    // ------------------------------------------------------------------
    // Per-element access (works for all types)
    // ------------------------------------------------------------------

    /// Read element at `index` as a WinRTValue.
    /// For Values arrays, returns a clone of the stored value.
    /// For CoTaskMem arrays, reads from raw bytes (AddRef / DuplicateString as needed).
    pub fn get(&self, index: usize) -> WinRTValue {
        assert!(index < self.len(), "ArrayData::get index {} out of bounds (len {})", index, self.len());
        match &self.buffer {
            ArrayBuffer::Values(v) => v[index].clone(),
            ArrayBuffer::CoTaskMem { ptr, .. } => {
                self.get_from_raw(index, *ptr as *const u8)
            }
        }
    }

    /// Read element from a raw byte buffer (CoTaskMem path).
    fn get_from_raw(&self, index: usize, base: *const u8) -> WinRTValue {
        let elem_size = self.element_type.element_size();
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
        match &self.buffer {
            ArrayBuffer::Values(v) => v[index].as_i32().unwrap(),
            ArrayBuffer::CoTaskMem { ptr, len } => {
                assert!(index < *len);
                unsafe { *((*ptr as *const u8).add(index * 4) as *const i32) }
            }
        }
    }

    // ------------------------------------------------------------------
    // ABI serialization (for PassArray FFI calls)
    // ------------------------------------------------------------------

    /// Serialize elements to a contiguous byte buffer for PassArray ABI.
    /// Returns an owned Vec<u8> that must be kept alive for the duration of the FFI call.
    pub(crate) fn serialize_for_abi(&self) -> Vec<u8> {
        match &self.buffer {
            ArrayBuffer::Values(values) => serialize_to_buffer(&self.element_type, values),
            ArrayBuffer::CoTaskMem { ptr, len } => {
                let elem_size = self.element_type.element_size();
                let total = *len * elem_size;
                let mut buf = vec![0u8; total];
                if total > 0 {
                    unsafe {
                        std::ptr::copy_nonoverlapping(*ptr as *const u8, buf.as_mut_ptr(), total);
                    }
                }
                buf
            }
        }
    }
}

// ======================================================================
// Drop — release non-blittable elements, then free the buffer
// ======================================================================

impl Drop for ArrayData {
    fn drop(&mut self) {
        // Values: Vec<WinRTValue> drops automatically, WinRTValue handles Release/DeleteString.
        // We only need manual cleanup for CoTaskMem.
        let buffer = std::mem::replace(
            &mut self.buffer,
            ArrayBuffer::CoTaskMem { ptr: std::ptr::null_mut(), len: 0 },
        );

        if let ArrayBuffer::CoTaskMem { ptr, len } = buffer {
            if len > 0 && !ptr.is_null() {
                let base = ptr as *const u8;
                let elem_size = self.element_type.element_size();
                let kind = self.element_type.kind();

                // Release non-blittable elements
                match kind {
                    TypeKind::HString => {
                        for i in 0..len {
                            unsafe {
                                let raw = *(base.add(i * elem_size) as *const *mut c_void);
                                if !raw.is_null() {
                                    let _hstr: windows_core::HSTRING = std::mem::transmute(raw);
                                }
                            }
                        }
                    }
                    kind if kind.is_com_pointer() => {
                        for i in 0..len {
                            unsafe {
                                let raw = *(base.add(i * elem_size) as *const *mut c_void);
                                if !raw.is_null() {
                                    let _obj = IUnknown::from_raw(raw);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }

            if !ptr.is_null() {
                unsafe {
                    windows::Win32::System::Com::CoTaskMemFree(Some(ptr));
                }
            }
        }
        // ArrayBuffer::Values is dropped automatically here
    }
}

// ======================================================================
// Clone
// ======================================================================

impl Clone for ArrayData {
    fn clone(&self) -> Self {
        match &self.buffer {
            ArrayBuffer::Values(v) => ArrayData {
                element_type: self.element_type.clone(),
                buffer: ArrayBuffer::Values(v.clone()),
            },
            ArrayBuffer::CoTaskMem { ptr, len } => {
                if *len == 0 || ptr.is_null() {
                    return ArrayData::empty(self.element_type.clone());
                }

                let elem_size = self.element_type.element_size();
                let total_bytes = *len * elem_size;
                let base = *ptr as *const u8;

                let new_ptr = unsafe {
                    windows::Win32::System::Com::CoTaskMemAlloc(total_bytes)
                };
                assert!(!new_ptr.is_null(), "CoTaskMemAlloc failed in ArrayData::clone");
                let new_buf = new_ptr as *mut u8;

                let kind = self.element_type.kind();
                match kind {
                    TypeKind::HString => {
                        unsafe { std::ptr::write_bytes(new_buf, 0, total_bytes) };
                        for i in 0..*len {
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
                        for i in 0..*len {
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
                        unsafe { std::ptr::copy_nonoverlapping(base, new_buf, total_bytes) };
                    }
                }

                ArrayData {
                    element_type: self.element_type.clone(),
                    buffer: ArrayBuffer::CoTaskMem { ptr: new_ptr, len: *len },
                }
            }
        }
    }
}

// ======================================================================
// Serialization — WinRTValue → raw bytes for PassArray ABI
// ======================================================================

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
