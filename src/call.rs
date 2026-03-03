use core::ffi::c_void;
use libffi::middle::{Arg, Cif, arg};
use windows_core::{HRESULT, HSTRING, IUnknown, Interface};

use crate::{abi::AbiValue, signature::Parameter, value::WinRTValue};

pub fn get_vtable_function_ptr(obj: *mut c_void, method_index: usize) -> *mut c_void {
    unsafe {
        // a function pointer is *const c_void (void* in C)
        // a vtable is a array of function pointers: *const *const c_void (void** in C)
        // a COM object pointer is a pointer to vtable: *const *const *const c_void (void*** in C)
        let vtable_ptr = *(obj as *const *const *mut c_void);
        *vtable_ptr.add(method_index)
    }
}

pub fn call_winrt_method_1<T1>(vtable_index: usize, obj: *mut c_void, x1: T1) -> HRESULT {
    let ptr : *mut c_void = unsafe { std::mem::transmute(&x1) };
    println!("Calling winrt method 1 vtable index: {} with obj {} x1 {}", vtable_index, obj as usize, unsafe { *(ptr as *mut usize) }); 
    let method_ptr = get_vtable_function_ptr(obj, vtable_index);

    unsafe {
        let method: extern "system" fn(*mut c_void, T1) -> HRESULT =
            std::mem::transmute(method_ptr);
        method(obj, x1)
    }
}

pub fn call_winrt_method_2<T1, T2>(
    vtable_index: usize,
    obj: *mut c_void,
    x1: T1,
    x2: T2,
) -> HRESULT {
    println!("Calling winrt method 2 vtable index: {}", vtable_index);
    let method_ptr = get_vtable_function_ptr(obj, vtable_index);

    unsafe {
        let method: extern "system" fn(*mut c_void, T1, T2) -> HRESULT =
            std::mem::transmute(method_ptr);
        method(obj, x1, x2)
    }
}

pub fn call_winrt_method_3<T1, T2, T3>(
    vtable_index: usize,
    obj: *mut c_void,
    x1: T1,
    x2: T2,
    x3: T3,
) -> HRESULT {
    println!("Calling winrt method 2 vtable index: {}", vtable_index);
    let method_ptr = get_vtable_function_ptr(obj, vtable_index);

    unsafe {
        let method: extern "system" fn(*mut c_void, T1, T2, T3) -> HRESULT =
            std::mem::transmute(method_ptr);
        method(obj, x1, x2, x3)
    }
}

use crate::array::{serialize_array_elements, deserialize_array_elements};

/// Stable heap storage for array in-param data (must outlive ffi call).
struct ArrayInSlot {
    length: u32,
    buffer: Vec<u8>,
    data_ptr: *const u8, // cached buffer.as_ptr(), stored for stable arg reference
}

/// Stable heap storage for array out-param data (callee writes into these fields).
struct ArrayOutSlot {
    length: u32,
    data_ptr: *mut c_void,
    element_type: crate::types::WinRTType,
}

/// Stable heap storage for FillArray out-param data (caller-allocated buffer).
struct FillArraySlot {
    capacity: u32,
    buffer: Vec<u8>,
    buffer_ptr: *mut u8, // cached buffer.as_mut_ptr()
    element_type: crate::types::WinRTType,
}

pub fn call_winrt_method_dynamic(
    vtable_index: usize,
    obj: *mut c_void,
    parameters: &[Parameter],
    args: &[WinRTValue],
    out_count: usize,
    cif: &libffi::middle::Cif,
) -> windows_core::Result<Vec<WinRTValue>> {
    use crate::type_table::ValueTypeData;
    use crate::types::WinRTType;
    use libffi::middle::CodePtr;

    let fptr = get_vtable_function_ptr(obj, vtable_index);
    let mut ffi_args: Vec<Arg> = Vec::with_capacity(parameters.len() * 2 + 1);
    let mut out_values: Vec<AbiValue> = Vec::with_capacity(out_count);
    let mut out_ptrs: Vec<*const std::ffi::c_void> = Vec::with_capacity(out_count);
    let mut struct_out_values: Vec<Option<ValueTypeData>> = Vec::with_capacity(out_count);

    // Array storage: Box'd for pointer stability (addresses don't change after creation)
    let mut array_out_slots: Vec<Box<ArrayOutSlot>> = Vec::new();
    // Map out value_index → array_out_slots index (None if not array)
    let mut array_out_map: Vec<Option<usize>> = Vec::with_capacity(out_count);
    // Pre-computed pointers into array_out_slots for use as ffi args
    let mut array_out_len_ptrs: Vec<*mut u32> = Vec::new();
    let mut array_out_data_ptrs: Vec<*mut *mut c_void> = Vec::new();

    // Array in-param storage: pre-compute all before building ffi_args
    let mut array_in_slots: Vec<Box<ArrayInSlot>> = Vec::new();

    // FillArray storage: caller-allocated buffers
    let mut fill_array_slots: Vec<Box<FillArraySlot>> = Vec::new();
    let mut fill_array_map: Vec<Option<usize>> = Vec::with_capacity(out_count);

    ffi_args.push(arg(&obj));

    // Phase 1a: Pre-allocate all out parameters
    for p in parameters {
        if p.is_out {
            if p.typ.is_fill_array() {
                // FillArray: caller allocates buffer. Use the capacity from args.
                let array_data = args[p.value_index]
                    .as_array()
                    .expect("Expected WinRTValue::Array with capacity for FillArray parameter");
                let elem_type = p.typ.array_element_type().clone();
                let capacity = array_data.len() as u32;
                let elem_size = elem_type.element_size();
                let mut buffer = vec![0u8; capacity as usize * elem_size];
                let buffer_ptr = buffer.as_mut_ptr();
                let slot = Box::new(FillArraySlot {
                    capacity,
                    buffer,
                    buffer_ptr,
                    element_type: elem_type,
                });
                fill_array_map.push(Some(fill_array_slots.len()));
                fill_array_slots.push(slot);
                // Placeholders for index alignment
                out_values.push(AbiValue::Pointer(std::ptr::null_mut()));
                out_ptrs.push(std::ptr::null());
                struct_out_values.push(None);
                array_out_map.push(None);
            } else if p.typ.is_array() {
                let slot = Box::new(ArrayOutSlot {
                    length: 0u32,
                    data_ptr: std::ptr::null_mut(),
                    element_type: p.typ.array_element_type().clone(),
                });
                let slot_idx = array_out_slots.len();
                array_out_map.push(Some(slot_idx));
                array_out_slots.push(slot);
                let slot_ref = &mut *array_out_slots[slot_idx];
                array_out_len_ptrs.push(&mut slot_ref.length);
                array_out_data_ptrs.push(&mut slot_ref.data_ptr);
                out_values.push(AbiValue::Pointer(std::ptr::null_mut()));
                out_ptrs.push(std::ptr::null());
                struct_out_values.push(None);
                fill_array_map.push(None);
            } else if let WinRTType::Struct(handle) = &p.typ {
                let val = handle.default_value();
                out_ptrs.push(val.as_ptr() as *const std::ffi::c_void);
                out_values.push(AbiValue::Pointer(std::ptr::null_mut()));
                struct_out_values.push(Some(val));
                array_out_map.push(None);
                fill_array_map.push(None);
            } else {
                out_values.push(p.typ.abi_type().default_value());
                out_ptrs.push(out_values.last().unwrap().as_out_ptr());
                struct_out_values.push(None);
                array_out_map.push(None);
                fill_array_map.push(None);
            }
        }
    }

    // Phase 1b: Pre-compute all array in-param data (must happen before Phase 2)
    for p in parameters {
        if !p.is_out && p.typ.is_array() {
            let array_data = args[p.value_index]
                .as_array()
                .expect("Expected WinRTValue::Array for array in-parameter");
            let buffer = serialize_array_elements(array_data);
            let mut slot = Box::new(ArrayInSlot {
                length: array_data.len() as u32,
                data_ptr: std::ptr::null(),
                buffer,
            });
            slot.data_ptr = slot.buffer.as_ptr();
            array_in_slots.push(slot);
        }
    }

    // Phase 2: Build ffi_args
    let mut array_in_idx = 0usize;
    let mut array_out_idx = 0usize;
    for p in parameters {
        if p.is_out {
            if let Some(slot_idx) = fill_array_map[p.value_index] {
                // FillArray: push TWO args (capacity value, buffer pointer value)
                let slot = &*fill_array_slots[slot_idx];
                ffi_args.push(arg(&slot.capacity));
                ffi_args.push(arg(&slot.buffer_ptr));
            } else if array_out_map[p.value_index].is_some() {
                // ReceiveArray out: push TWO args (pointer-to-length, pointer-to-data_ptr)
                ffi_args.push(arg(&array_out_len_ptrs[array_out_idx]));
                ffi_args.push(arg(&array_out_data_ptrs[array_out_idx]));
                array_out_idx += 1;
            } else {
                ffi_args.push(arg(&out_ptrs[p.value_index]));
            }
        } else if p.typ.is_array() {
            // Array in: push TWO args (length value, data pointer value)
            let slot = &*array_in_slots[array_in_idx];
            ffi_args.push(arg(&slot.length));
            ffi_args.push(arg(&slot.data_ptr));
            array_in_idx += 1;
        } else {
            ffi_args.push(args[p.value_index].libffi_arg());
        }
    }

    // Phase 3: Call
    let hr: windows_core::HRESULT = unsafe { cif.call(CodePtr(fptr), &ffi_args) };
    hr.ok()?;

    // Phase 4: Extract results
    let mut result_values: Vec<WinRTValue> = Vec::with_capacity(out_count);
    for p in parameters {
        if p.is_out {
            if let Some(slot_idx) = fill_array_map[p.value_index] {
                // FillArray: deserialize from caller-allocated buffer (full capacity)
                let slot = &fill_array_slots[slot_idx];
                let capacity = slot.capacity as usize;
                let elements = deserialize_array_elements(
                    &slot.element_type,
                    slot.buffer_ptr as *mut c_void,
                    capacity,
                );
                result_values.push(WinRTValue::Array(
                    crate::value::ArrayData::from_values(slot.element_type.clone(), elements)
                ));
            } else if let Some(slot_idx) = array_out_map[p.value_index] {
                // Array out: read callee-written length + data_ptr, deserialize, free.
                // Use a guard to ensure CoTaskMemFree is called even if deserialize panics.
                let slot = &array_out_slots[slot_idx];
                let length = slot.length as usize;
                let data_ptr = slot.data_ptr;
                let array_value = if data_ptr.is_null() || length == 0 {
                    crate::value::ArrayData::empty(slot.element_type.clone())
                } else {
                    struct CoTaskMemGuard(*mut c_void);
                    impl Drop for CoTaskMemGuard {
                        fn drop(&mut self) {
                            unsafe { windows::Win32::System::Com::CoTaskMemFree(Some(self.0)); }
                        }
                    }
                    let _guard = CoTaskMemGuard(data_ptr);
                    let elements = deserialize_array_elements(&slot.element_type, data_ptr, length);
                    // _guard drops here (or on panic), freeing the callee-allocated buffer
                    crate::value::ArrayData::from_values(slot.element_type.clone(), elements)
                };
                result_values.push(WinRTValue::Array(array_value));
            } else if let Some(struct_val) = struct_out_values[p.value_index].take() {
                result_values.push(WinRTValue::Struct(struct_val));
            } else {
                let out_value = p.typ.from_out_value(&out_values[p.value_index]);
                result_values.push(out_value.unwrap());
            }
        }
    }
    Ok(result_values)
}
