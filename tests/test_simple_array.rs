// Simple test for PropertyValue.CreateInt32Array
// This is the SIMPLEST possible array API to test with

use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};
use windows_core::{GUID, HSTRING, IUnknown, Interface, h};

#[test]
fn test_property_value_create_int32_array_manual() -> windows::core::Result<()> {
    // Initialize WinRT
    unsafe { RoInitialize(RO_INIT_MULTITHREADED) }?;

    // Step 1: Get PropertyValue activation factory
    let factory = unsafe {
        windows::Win32::System::WinRT::RoGetActivationFactory::<windows::Win32::System::WinRT::IActivationFactory>(
            h!("Windows.Foundation.PropertyValue")
        )?
    };

    // Step 2: Cast to IPropertyValueStatics
    // GUID: {629BDBC8-D932-4FF4-96B9-8D96C5C1E858}
    let statics_guid = GUID::from_u128(0x629BDBC8_D932_4FF4_96B9_8D96C5C1E858);

    let mut statics_ptr = std::ptr::null_mut();
    unsafe { factory.query(&statics_guid, &mut statics_ptr) }.ok()?;
    let statics = unsafe { IUnknown::from_raw(statics_ptr) };

    // Step 3: Prepare test array
    let test_data = vec![1i32, 2, 3, 4, 5];
    let length = test_data.len() as u32;
    let data_ptr = test_data.as_ptr();

    println!("Test data: {:?}", test_data);
    println!("Length: {}, Ptr: {:p}", length, data_ptr);

    // Step 4: Call CreateInt32Array
    // Find the method index first using find_createint32array_index test
    // For now, we'll try to determine it

    // IPropertyValueStatics methods start after IInspectable (6 methods)
    // We need to find which index CreateInt32Array is at

    // Let's use static projection to verify the call works
    use windows::Foundation::PropertyValue;

    let prop_value = PropertyValue::CreateInt32Array(&test_data)?;
    println!("✓ Successfully created PropertyValue with static projection");
    println!("  Result: {:?}", prop_value);

    Ok(())
}

#[test]
fn test_property_value_create_int32_array_dynamic() -> windows::core::Result<()> {
    // Initialize WinRT
    unsafe { RoInitialize(RO_INIT_MULTITHREADED) }?;

    // Get factory and cast to statics
    let factory = unsafe {
        windows::Win32::System::WinRT::RoGetActivationFactory::<windows::Win32::System::WinRT::IActivationFactory>(
            h!("Windows.Foundation.PropertyValue")
        )?
    };

    let statics_guid = GUID::from_u128(0x629BDBC8_D932_4FF4_96B9_8D96C5C1E858);
    let mut statics_ptr = std::ptr::null_mut();
    unsafe { factory.query(&statics_guid, &mut statics_ptr) }.ok()?;
    let statics = unsafe { IUnknown::from_raw(statics_ptr) };

    // Prepare array
    let test_data = vec![10i32, 20, 30, 40, 50];
    let length = test_data.len() as u32;
    let data_ptr = test_data.as_ptr();

    // Manual dynamic call
    // Method signature: HRESULT CreateInt32Array(uint32_t length, int32_t* data, IInspectable** result)

    // Get vtable function pointer
    let vtable_index = 8; // PLACEHOLDER - need to find actual index

    let method_ptr = unsafe {
        let obj = statics.as_raw();
        let vtable_ptr = *(obj as *const *const *mut std::ffi::c_void);
        *vtable_ptr.add(vtable_index)
    };

    // Call the method
    let mut result: *mut std::ffi::c_void = std::ptr::null_mut();

    let hr: windows_core::HRESULT = unsafe {
        let method: extern "system" fn(
            *mut std::ffi::c_void,  // this
            u32,                     // length
            *const i32,              // data
            *mut *mut std::ffi::c_void, // out result
        ) -> windows_core::HRESULT = std::mem::transmute(method_ptr);

        method(statics.as_raw(), length, data_ptr, &mut result)
    };

    if hr.is_ok() {
        println!("✓ Dynamic call succeeded!");
        let result_inspectable = unsafe { windows_core::IInspectable::from_raw(result) };
        println!("  Result: {:?}", result_inspectable);
        Ok(())
    } else {
        println!("✗ Dynamic call failed: {:?}", hr);
        Err(windows::core::Error::from(hr))
    }
}

// Helper: Test just getting the factory
#[test]
fn test_get_property_value_factory() -> windows::core::Result<()> {
    unsafe { RoInitialize(RO_INIT_MULTITHREADED) }?;

    let factory = unsafe {
        windows::Win32::System::WinRT::RoGetActivationFactory::<windows::Win32::System::WinRT::IActivationFactory>(
            h!("Windows.Foundation.PropertyValue")
        )?
    };

    println!("✓ Got PropertyValue factory: {:?}", factory);

    // Cast to IPropertyValueStatics
    let statics_guid = GUID::from_u128(0x629BDBC8_D932_4FF4_96B9_8D96C5C1E858);
    let mut statics_ptr = std::ptr::null_mut();
    unsafe { factory.query(&statics_guid, &mut statics_ptr) }.ok()?;

    if !statics_ptr.is_null() {
        println!("✓ Successfully cast to IPropertyValueStatics");
        let statics = unsafe { IUnknown::from_raw(statics_ptr) };
        println!("  Statics interface: {:?}", statics);
    }

    Ok(())
}
