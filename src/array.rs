use core::ffi::c_void;
use windows_core::{IUnknown, Interface};

use crate::types::WinRTType;
use crate::value::WinRTValue;

/// Holds a dynamically-typed WinRT array for use with PassArray/ReceiveArray ABI patterns.
#[derive(Debug, Clone)]
pub struct ArrayData {
    pub element_type: WinRTType,
    pub elements: Vec<WinRTValue>,
}

impl ArrayData {
    pub fn empty(element_type: WinRTType) -> Self {
        ArrayData {
            element_type,
            elements: Vec::new(),
        }
    }

    pub fn from_values(element_type: WinRTType, elements: Vec<WinRTValue>) -> Self {
        ArrayData {
            element_type,
            elements,
        }
    }

    pub fn len(&self) -> usize {
        self.elements.len()
    }
}

/// Serialize array elements into a contiguous byte buffer for PassArray ABI.
pub(crate) fn serialize_array_elements(data: &ArrayData) -> Vec<u8> {
    let elem_size = data.element_type.element_size();
    let mut buffer = Vec::with_capacity(data.elements.len() * elem_size);
    for elem in &data.elements {
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
                // GUID is 16 bytes: {u32, u16, u16, [u8; 8]}
                let bytes: &[u8; 16] = unsafe { &*(g as *const windows_core::GUID as *const [u8; 16]) };
                buffer.extend_from_slice(bytes);
            }
            WinRTValue::Struct(vd) => {
                let size = vd.type_handle().size_of();
                let src = unsafe { std::slice::from_raw_parts(vd.as_ptr(), size) };
                buffer.extend_from_slice(src);
            }
            _ => panic!("Unsupported array element type for serialization: {:?}", elem.get_type()),
        }
    }
    buffer
}

/// Deserialize a raw buffer into WinRTValue elements for ReceiveArray ABI.
pub(crate) fn deserialize_array_elements(
    element_type: &WinRTType,
    data_ptr: *mut c_void,
    length: usize,
) -> Vec<WinRTValue> {
    let mut elements = Vec::with_capacity(length);
    unsafe {
        match element_type {
            WinRTType::Bool => {
                let ptr = data_ptr as *const u8;
                for i in 0..length { elements.push(WinRTValue::Bool(*ptr.add(i) != 0)); }
            }
            WinRTType::I8 => {
                let ptr = data_ptr as *const i8;
                for i in 0..length { elements.push(WinRTValue::I8(*ptr.add(i))); }
            }
            WinRTType::U8 => {
                let ptr = data_ptr as *const u8;
                for i in 0..length { elements.push(WinRTValue::U8(*ptr.add(i))); }
            }
            WinRTType::I16 => {
                let ptr = data_ptr as *const i16;
                for i in 0..length { elements.push(WinRTValue::I16(*ptr.add(i))); }
            }
            WinRTType::U16 => {
                let ptr = data_ptr as *const u16;
                for i in 0..length { elements.push(WinRTValue::U16(*ptr.add(i))); }
            }
            WinRTType::I32 => {
                let ptr = data_ptr as *const i32;
                for i in 0..length { elements.push(WinRTValue::I32(*ptr.add(i))); }
            }
            WinRTType::U32 => {
                let ptr = data_ptr as *const u32;
                for i in 0..length { elements.push(WinRTValue::U32(*ptr.add(i))); }
            }
            WinRTType::I64 => {
                let ptr = data_ptr as *const i64;
                for i in 0..length { elements.push(WinRTValue::I64(*ptr.add(i))); }
            }
            WinRTType::U64 => {
                let ptr = data_ptr as *const u64;
                for i in 0..length { elements.push(WinRTValue::U64(*ptr.add(i))); }
            }
            WinRTType::F32 => {
                let ptr = data_ptr as *const f32;
                for i in 0..length { elements.push(WinRTValue::F32(*ptr.add(i))); }
            }
            WinRTType::F64 => {
                let ptr = data_ptr as *const f64;
                for i in 0..length { elements.push(WinRTValue::F64(*ptr.add(i))); }
            }
            WinRTType::Char16 => {
                let ptr = data_ptr as *const u16;
                for i in 0..length { elements.push(WinRTValue::U16(*ptr.add(i))); }
            }
            WinRTType::Guid => {
                // GUID is 16 bytes, blittable
                let base = data_ptr as *const u8;
                for i in 0..length {
                    let guid = *(base.add(i * 16) as *const windows_core::GUID);
                    elements.push(WinRTValue::Guid(guid));
                }
            }
            WinRTType::HString => {
                // Each element is an HSTRING handle (pointer-sized).
                // ReceiveArray: callee gives ownership of each handle to the caller.
                let ptr = data_ptr as *const *mut c_void;
                for i in 0..length {
                    let raw = *ptr.add(i);
                    let hstr: windows_core::HSTRING = std::mem::transmute(raw);
                    elements.push(WinRTValue::HString(hstr));
                }
            }
            WinRTType::Object | WinRTType::Interface(_) => {
                let ptr = data_ptr as *const *mut c_void;
                for i in 0..length {
                    let raw = *ptr.add(i);
                    elements.push(WinRTValue::Object(IUnknown::from_raw(raw)));
                }
            }
            WinRTType::Struct(handle) => {
                // Struct elements are contiguous blittable data
                let elem_size = handle.size_of();
                let base = data_ptr as *const u8;
                for i in 0..length {
                    let mut vd = handle.default_value();
                    std::ptr::copy_nonoverlapping(
                        base.add(i * elem_size),
                        vd.as_mut_ptr(),
                        elem_size,
                    );
                    elements.push(WinRTValue::Struct(vd));
                }
            }
            _ => panic!("Unsupported array element type for deserialization: {:?}", element_type),
        }
    }
    elements
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
        // let buffer = CryptographicBuffer::GenerateRandom(128)?;
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
        // get current device location
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
                    // f64,
                    // f64,
                    // f64,
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
        let t : libffi::low::ffi_type = unsafe { std::mem::zeroed() };

        let BasicGeoPositionStruct = libffi::middle::Type::structure(vec![
            libffi::middle::Type::f64(), // Latitude
            libffi::middle::Type::f64(), // Longitude
            libffi::middle::Type::f64(), // Altitude
        ]);
        let f1 = Layout::new::<f64>();
        let (f2, offset2) = f1.extend(Layout::new::<f64>()).unwrap();
        let (f3, offset3) = f2.extend(Layout::new::<f64>()).unwrap();
        let sl = f3.pad_to_align();
        println!("Struct layout size: {}, align: {}", sl.size(), sl.align());
        println!("Field offsets: f1: 0, f2: {}, f3: {}", offset2, offset3);

        // let sptr = unsafe { (&position as *const _ as *mut std::ffi::c_void).add(0) };
        let sptr = unsafe { std::alloc::alloc(sl) };
        let pf1 = unsafe { sptr } as *mut f64;
        let pf2 = unsafe { sptr.add(offset2) } as *mut f64;
        let pf3 = unsafe { sptr.add(offset3) } as *mut f64;
        unsafe {
            *pf1 = 11.0;
            *pf2 = 22.0;
            *pf3 = 33.0;
        }
        println!(
            "After modifying fields {}, {}, {} ",
            position.Latitude, position.Longitude, position.Altitude
        );
        println!(
            "Struct values: f1: {}, f2: {}, f3: {}",
            unsafe { *pf1 },
            unsafe { *pf2 },
            unsafe { *pf3 }
        );

        println!(
            "position size : {} , align: {}, ptr: {:?}",
            std::mem::size_of::<BasicGeoposition>(),
            std::mem::align_of::<BasicGeoposition>(),
            &position as *const _
        );

        println!("dynamic struct ptr: {:?}", sptr);

        let create = libffi::middle::Cif::new(
            vec![
                libffi::middle::Type::pointer(), // this pointer
                BasicGeoPositionStruct,          // BasicGeoposition
                libffi::middle::Type::pointer(), // out parameter
            ]
            .into_iter(),
            libffi::middle::Type::i32(), // HRESULT
        );
        let GeopointFactory = afactory.cast::<windows::Devices::Geolocation::IGeopointFactory>()?;
        let createFptr = get_vtable_function_ptr(GeopointFactory.as_raw(), 6);
        let mut out = std::ptr::null_mut();
        let pOut = &mut out as *mut *mut std::ffi::c_void;
        let thisPtr = GeopointFactory.as_raw();
        let hr = unsafe {
            libffi::low::call::<HRESULT>(
                create.as_raw_ptr(),
                CodePtr(createFptr),
                vec![
                    &GeopointFactory.as_raw() as *const _ as *mut std::ffi::c_void,
                    // &position as *const _ as *mut std::ffi::c_void,
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
