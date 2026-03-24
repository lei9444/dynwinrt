#![allow(unsafe_op_in_unsafe_fn)]
//! Dynamic WinRT IVector<T> / IIterable<T> / IVectorView<T> / IIterator<T> implementation.
//!
//! Creates COM objects at runtime that implement the WinRT collection interfaces,
//! allowing JS callers to construct vectors and pass them to WinRT APIs.

use core::ffi::c_void;
use std::cell::RefCell;
use windows_core::{GUID, HRESULT, IUnknown, Interface};

use crate::com_helpers::{
    IInspectableVtbl, E_BOUNDS, S_OK,
    com_to_usize, com_usize_addref_out, com_usize_release,
};
use crate::com_helpers::{inspectable_stubs, dual_vtable_com, single_vtable_com, impl_drop_release_items};

// ======================================================================
// IIDs for collection PIIDs
// ======================================================================

/// All IIDs needed for an IVector<T> collection.
#[derive(Debug, Clone)]
pub struct VectorIids {
    pub iterable: GUID,
    pub vector: GUID,
    pub vector_view: GUID,
    pub iterator: GUID,
}

// ======================================================================
// COM vtable layouts (matching WinRT ABI)
// ======================================================================

/// IIterable<T> vtable: IInspectable + First()
#[repr(C)]
struct IterableVtbl {
    base: IInspectableVtbl,
    first: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> HRESULT,
}

/// IVector<T> vtable: IInspectable + 12 methods
#[repr(C)]
struct VectorVtbl {
    base: IInspectableVtbl,
    get_at: unsafe extern "system" fn(*mut c_void, u32, *mut *mut c_void) -> HRESULT,
    get_size: unsafe extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
    get_view: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> HRESULT,
    index_of: unsafe extern "system" fn(*mut c_void, *mut c_void, *mut u32, *mut bool) -> HRESULT,
    set_at: unsafe extern "system" fn(*mut c_void, u32, *mut c_void) -> HRESULT,
    insert_at: unsafe extern "system" fn(*mut c_void, u32, *mut c_void) -> HRESULT,
    remove_at: unsafe extern "system" fn(*mut c_void, u32) -> HRESULT,
    append: unsafe extern "system" fn(*mut c_void, *mut c_void) -> HRESULT,
    remove_at_end: unsafe extern "system" fn(*mut c_void) -> HRESULT,
    clear: unsafe extern "system" fn(*mut c_void) -> HRESULT,
    get_many: unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void, *mut u32) -> HRESULT,
    replace_all: unsafe extern "system" fn(*mut c_void, u32, *const *mut c_void) -> HRESULT,
}

/// IVectorView<T> vtable: IInspectable + 4 methods
#[repr(C)]
struct VectorViewVtbl {
    base: IInspectableVtbl,
    get_at: unsafe extern "system" fn(*mut c_void, u32, *mut *mut c_void) -> HRESULT,
    get_size: unsafe extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
    index_of: unsafe extern "system" fn(*mut c_void, *mut c_void, *mut u32, *mut bool) -> HRESULT,
    get_many: unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void, *mut u32) -> HRESULT,
}

/// IIterator<T> vtable: IInspectable + 4 methods
#[repr(C)]
struct IteratorVtbl {
    base: IInspectableVtbl,
    get_current: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> HRESULT,
    get_has_current: unsafe extern "system" fn(*mut c_void, *mut bool) -> HRESULT,
    move_next: unsafe extern "system" fn(*mut c_void, *mut bool) -> HRESULT,
    get_many: unsafe extern "system" fn(*mut c_void, u32, *mut *mut c_void, *mut u32) -> HRESULT,
}


/// Write a raw usize item to an output pointer, AddRef'ing if it's a COM reference type.
#[inline(always)]
unsafe fn write_item_out(is_value_type: bool, raw: usize, result: *mut *mut c_void) {
    if is_value_type {
        *(result as *mut usize) = raw;
    } else {
        *result = com_usize_addref_out(raw);
    }
}


// ======================================================================
// SingleThreadedVector
// ======================================================================

/// A dynamically-constructed WinRT IVector<T> + IIterable<T> COM object.
///
/// Stores items as raw `usize` values. For reference types (COM objects),
/// each usize is a raw IUnknown pointer with manual AddRef/Release.
/// For value types (structs ≤ pointer size), each usize holds the struct
/// bytes directly — no refcounting needed.
#[repr(C)]
struct SingleThreadedVector {
    vtable_iterable: *const IterableVtbl,
    vtable_vector: *const VectorVtbl,
    ref_count: windows_core::imp::RefCount,
    items: RefCell<Vec<usize>>,
    is_value_type: bool,
    iids: VectorIids,
}

unsafe impl Send for SingleThreadedVector {}
unsafe impl Sync for SingleThreadedVector {}

impl SingleThreadedVector {
    const ITERABLE_VTBL: IterableVtbl = IterableVtbl {
        base: IInspectableVtbl {
            base: windows_core::IUnknown_Vtbl {
                QueryInterface: Self::qi_iterable,
                AddRef: Self::add_ref_iterable,
                Release: Self::release_iterable,
            },
            get_iids: Self::get_iids_iterable,
            get_runtime_class_name: Self::get_runtime_class_name_iterable,
            get_trust_level: Self::get_trust_level_iterable,
        },
        first: Self::first,
    };

    const VECTOR_VTBL: VectorVtbl = VectorVtbl {
        base: IInspectableVtbl {
            base: windows_core::IUnknown_Vtbl {
                QueryInterface: Self::qi_vector,
                AddRef: Self::add_ref_vector,
                Release: Self::release_vector,
            },
            get_iids: Self::get_iids_vector,
            get_runtime_class_name: Self::get_runtime_class_name_vector,
            get_trust_level: Self::get_trust_level_vector,
        },
        get_at: Self::get_at,
        get_size: Self::get_size,
        get_view: Self::get_view,
        index_of: Self::index_of,
        set_at: Self::set_at,
        insert_at: Self::insert_at,
        remove_at: Self::remove_at,
        append: Self::append,
        remove_at_end: Self::remove_at_end,
        clear: Self::clear,
        get_many: Self::get_many,
        replace_all: Self::replace_all,
    };

    dual_vtable_com!(iterable, vector, vector);
    inspectable_stubs!(iterable, vector);

    // ------------------------------------------------------------------
    // IIterable<T>
    // ------------------------------------------------------------------

    unsafe extern "system" fn first(
        this: *mut c_void,
        result: *mut *mut c_void,
    ) -> HRESULT {
        let me = Self::from_iterable_ptr(this);
        let items = me.items.borrow();
        let snapshot = if me.is_value_type {
            items.clone()
        } else {
            items.iter().map(|&raw| com_to_usize(raw as *mut c_void)).collect()
        };
        let iter = SingleThreadedIterator::create(snapshot, me.is_value_type, me.iids.iterator);
        *result = iter.into_raw();
        S_OK
    }

    // ------------------------------------------------------------------
    // IVector<T>
    // ------------------------------------------------------------------

    unsafe extern "system" fn get_at(
        this: *mut c_void,
        index: u32,
        result: *mut *mut c_void,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let items = me.items.borrow();
        if (index as usize) >= items.len() {
            return E_BOUNDS;
        }
        let raw = items[index as usize];
        write_item_out(me.is_value_type, raw, result);
        S_OK
    }

    unsafe extern "system" fn get_size(
        this: *mut c_void,
        result: *mut u32,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        *result = me.items.borrow().len() as u32;
        S_OK
    }

    unsafe extern "system" fn get_view(
        this: *mut c_void,
        result: *mut *mut c_void,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let items = me.items.borrow();
        let snapshot = if me.is_value_type {
            items.clone()
        } else {
            items.iter().map(|&raw| com_to_usize(raw as *mut c_void)).collect()
        };
        let view = SingleThreadedVectorView::create(snapshot, me.is_value_type, me.iids.clone());
        *result = view.into_raw();
        S_OK
    }

    unsafe extern "system" fn index_of(
        this: *mut c_void,
        value: *mut c_void,
        index: *mut u32,
        found: *mut bool,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let items = me.items.borrow();
        let needle = value as usize;
        for (i, &item) in items.iter().enumerate() {
            if item == needle {
                *index = i as u32;
                *found = true;
                return S_OK;
            }
        }
        *index = 0;
        *found = false;
        S_OK
    }

    unsafe extern "system" fn set_at(
        this: *mut c_void,
        index: u32,
        value: *mut c_void,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let mut items = me.items.borrow_mut();
        if (index as usize) >= items.len() {
            return E_BOUNDS;
        }
        let old = items[index as usize];
        items[index as usize] = if me.is_value_type {
            value as usize
        } else {
            let new_val = com_to_usize(value);
            com_usize_release(old);
            new_val
        };
        S_OK
    }

    unsafe extern "system" fn insert_at(
        this: *mut c_void,
        index: u32,
        value: *mut c_void,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let mut items = me.items.borrow_mut();
        if (index as usize) > items.len() {
            return E_BOUNDS;
        }
        let val = if me.is_value_type { value as usize } else { com_to_usize(value) };
        items.insert(index as usize, val);
        S_OK
    }

    unsafe extern "system" fn remove_at(
        this: *mut c_void,
        index: u32,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let mut items = me.items.borrow_mut();
        if (index as usize) >= items.len() {
            return E_BOUNDS;
        }
        let removed = items.remove(index as usize);
        if !me.is_value_type { com_usize_release(removed); }
        S_OK
    }

    unsafe extern "system" fn append(
        this: *mut c_void,
        value: *mut c_void,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let val = if me.is_value_type { value as usize } else { com_to_usize(value) };
        me.items.borrow_mut().push(val);
        S_OK
    }

    unsafe extern "system" fn remove_at_end(this: *mut c_void) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let mut items = me.items.borrow_mut();
        if items.is_empty() {
            return E_BOUNDS;
        }
        let removed = items.pop().unwrap();
        if !me.is_value_type { com_usize_release(removed); }
        S_OK
    }

    unsafe extern "system" fn clear(this: *mut c_void) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let old_items: Vec<usize> = me.items.borrow_mut().drain(..).collect();
        if !me.is_value_type {
            for raw in old_items { com_usize_release(raw); }
        }
        S_OK
    }

    unsafe extern "system" fn get_many(
        this: *mut c_void,
        start_index: u32,
        capacity: u32,
        items_out: *mut *mut c_void,
        actual: *mut u32,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let items = me.items.borrow();
        let start = start_index as usize;
        if start > items.len() {
            *actual = 0;
            return S_OK;
        }
        let count = std::cmp::min(capacity as usize, items.len() - start);
        for i in 0..count {
            let raw = items[start + i];
            write_item_out(me.is_value_type, raw, items_out.add(i));
        }
        *actual = count as u32;
        S_OK
    }

    unsafe extern "system" fn replace_all(
        this: *mut c_void,
        count: u32,
        values: *const *mut c_void,
    ) -> HRESULT {
        let me = Self::from_vector_ptr(this);
        let old_items: Vec<usize> = me.items.borrow_mut().drain(..).collect();
        if !me.is_value_type {
            for raw in old_items { com_usize_release(raw); }
        }
        let mut items = me.items.borrow_mut();
        for i in 0..count as usize {
            let raw = *values.add(i);
            let val = if me.is_value_type { raw as usize } else { com_to_usize(raw) };
            items.push(val);
        }
        S_OK
    }
}

impl_drop_release_items!(SingleThreadedVector, borrow);

// ======================================================================
// SingleThreadedVectorView
// ======================================================================

#[repr(C)]
struct SingleThreadedVectorView {
    vtable_iterable: *const IterableVtbl,
    vtable_view: *const VectorViewVtbl,
    ref_count: windows_core::imp::RefCount,
    items: Vec<usize>,
    is_value_type: bool,
    iids: VectorIids,
}

unsafe impl Send for SingleThreadedVectorView {}
unsafe impl Sync for SingleThreadedVectorView {}

impl SingleThreadedVectorView {
    const ITERABLE_VTBL: IterableVtbl = IterableVtbl {
        base: IInspectableVtbl {
            base: windows_core::IUnknown_Vtbl {
                QueryInterface: Self::qi_iterable,
                AddRef: Self::add_ref_iterable,
                Release: Self::release_iterable,
            },
            get_iids: Self::get_iids_stub,
            get_runtime_class_name: Self::get_runtime_class_name_stub,
            get_trust_level: Self::get_trust_level_stub,
        },
        first: Self::first,
    };

    const VIEW_VTBL: VectorViewVtbl = VectorViewVtbl {
        base: IInspectableVtbl {
            base: windows_core::IUnknown_Vtbl {
                QueryInterface: Self::qi_view,
                AddRef: Self::add_ref_view,
                Release: Self::release_view,
            },
            get_iids: Self::get_iids_stub2,
            get_runtime_class_name: Self::get_runtime_class_name_stub2,
            get_trust_level: Self::get_trust_level_stub2,
        },
        get_at: Self::get_at,
        get_size: Self::get_size,
        index_of: Self::index_of,
        get_many: Self::get_many,
    };

    fn create(items: Vec<usize>, is_value_type: bool, iids: VectorIids) -> IUnknown {
        let view = Box::new(Self {
            vtable_iterable: &Self::ITERABLE_VTBL,
            vtable_view: &Self::VIEW_VTBL,
            ref_count: windows_core::imp::RefCount::new(1),
            items,
            is_value_type,
            iids,
        });
        unsafe { IUnknown::from_raw(Box::into_raw(view) as *mut c_void) }
    }

    dual_vtable_com!(iterable, view, vector_view);
    inspectable_stubs!(stub, stub2);

    // -- IIterable<T> --

    unsafe extern "system" fn first(this: *mut c_void, result: *mut *mut c_void) -> HRESULT {
        let me = Self::from_iterable_ptr(this);
        let snapshot = if me.is_value_type {
            me.items.clone()
        } else {
            me.items.iter().map(|&raw| com_to_usize(raw as *mut c_void)).collect()
        };
        let iter = SingleThreadedIterator::create(snapshot, me.is_value_type, me.iids.iterator);
        *result = iter.into_raw();
        S_OK
    }

    // -- IVectorView<T> --

    unsafe extern "system" fn get_at(this: *mut c_void, index: u32, result: *mut *mut c_void) -> HRESULT {
        let me = Self::from_view_ptr(this);
        if (index as usize) >= me.items.len() { return E_BOUNDS; }
        let raw = me.items[index as usize];
        write_item_out(me.is_value_type, raw, result);
        S_OK
    }

    unsafe extern "system" fn get_size(this: *mut c_void, result: *mut u32) -> HRESULT {
        let me = Self::from_view_ptr(this);
        *result = me.items.len() as u32;
        S_OK
    }

    unsafe extern "system" fn index_of(this: *mut c_void, value: *mut c_void, index: *mut u32, found: *mut bool) -> HRESULT {
        let me = Self::from_view_ptr(this);
        let needle = value as usize;
        for (i, &item) in me.items.iter().enumerate() {
            if item == needle {
                *index = i as u32;
                *found = true;
                return S_OK;
            }
        }
        *index = 0;
        *found = false;
        S_OK
    }

    unsafe extern "system" fn get_many(this: *mut c_void, start_index: u32, capacity: u32, items_out: *mut *mut c_void, actual: *mut u32) -> HRESULT {
        let me = Self::from_view_ptr(this);
        let start = start_index as usize;
        if start > me.items.len() {
            *actual = 0;
            return S_OK;
        }
        let count = std::cmp::min(capacity as usize, me.items.len() - start);
        for i in 0..count {
            let raw = me.items[start + i];
            write_item_out(me.is_value_type, raw, items_out.add(i));
        }
        *actual = count as u32;
        S_OK
    }
}

impl_drop_release_items!(SingleThreadedVectorView, direct);

// ======================================================================
// SingleThreadedIterator
// ======================================================================

#[repr(C)]
pub(crate) struct SingleThreadedIterator {
    vtable: *const IteratorVtbl,
    ref_count: windows_core::imp::RefCount,
    items: Vec<usize>,
    is_value_type: bool,
    cursor: RefCell<usize>,
    iid_iterator: GUID,
}

unsafe impl Send for SingleThreadedIterator {}
unsafe impl Sync for SingleThreadedIterator {}

impl SingleThreadedIterator {
    const VTBL: IteratorVtbl = IteratorVtbl {
        base: IInspectableVtbl {
            base: windows_core::IUnknown_Vtbl {
                QueryInterface: Self::qi,
                AddRef: Self::add_ref,
                Release: Self::release,
            },
            get_iids: Self::get_iids_stub,
            get_runtime_class_name: Self::get_runtime_class_name_stub,
            get_trust_level: Self::get_trust_level_stub,
        },
        get_current: Self::get_current,
        get_has_current: Self::get_has_current,
        move_next: Self::move_next,
        get_many: Self::get_many,
    };

    pub(crate) fn create(items: Vec<usize>, is_value_type: bool, iid_iterator: GUID) -> IUnknown {
        let iter = Box::new(Self {
            vtable: &Self::VTBL,
            ref_count: windows_core::imp::RefCount::new(1),
            items,
            is_value_type,
            cursor: RefCell::new(0),
            iid_iterator,
        });
        unsafe { IUnknown::from_raw(Box::into_raw(iter) as *mut c_void) }
    }

    single_vtable_com!(|me: &Self| me.iid_iterator);
    inspectable_stubs!(stub);

    unsafe extern "system" fn get_current(this: *mut c_void, result: *mut *mut c_void) -> HRESULT {
        let me = &*(this as *const Self);
        let cursor = *me.cursor.borrow();
        if cursor >= me.items.len() { return E_BOUNDS; }
        let raw = me.items[cursor];
        write_item_out(me.is_value_type, raw, result);
        S_OK
    }

    unsafe extern "system" fn get_has_current(this: *mut c_void, result: *mut bool) -> HRESULT {
        let me = &*(this as *const Self);
        *result = *me.cursor.borrow() < me.items.len();
        S_OK
    }

    unsafe extern "system" fn move_next(this: *mut c_void, result: *mut bool) -> HRESULT {
        let me = &*(this as *const Self);
        let mut cursor = me.cursor.borrow_mut();
        if *cursor < me.items.len() {
            *cursor += 1;
        }
        *result = *cursor < me.items.len();
        S_OK
    }

    unsafe extern "system" fn get_many(this: *mut c_void, capacity: u32, items_out: *mut *mut c_void, actual: *mut u32) -> HRESULT {
        let me = &*(this as *const Self);
        let mut cursor = me.cursor.borrow_mut();
        let remaining = me.items.len().saturating_sub(*cursor);
        let count = std::cmp::min(capacity as usize, remaining);
        for i in 0..count {
            let raw = me.items[*cursor + i];
            write_item_out(me.is_value_type, raw, items_out.add(i));
        }
        *cursor += count;
        *actual = count as u32;
        S_OK
    }
}

impl_drop_release_items!(SingleThreadedIterator, direct);

// ======================================================================
// Public API
// ======================================================================

/// Create an IVector<T> COM object from WinRTValue items.
///
/// Automatically handles both reference types (COM objects → AddRef/Release)
/// and value types (structs ≤ pointer size → raw bytes, no refcounting).
pub fn create_vector_from_values(
    items: &[crate::WinRTValue],
    is_value_type: bool,
    elem_size: usize,
    iids: VectorIids,
) -> IUnknown {
    let packed = if is_value_type {
        assert!(
            items.is_empty() || elem_size <= std::mem::size_of::<usize>(),
            "create_vector: struct elem_size {} exceeds pointer size; not yet supported",
            elem_size
        );
        items.iter().map(|item| {
            let data = item.as_struct().expect("struct-typed vector requires Struct values");
            let mut val: usize = 0;
            unsafe {
                std::ptr::copy_nonoverlapping(
                    data.as_ptr(),
                    &mut val as *mut usize as *mut u8,
                    elem_size.min(std::mem::size_of::<usize>()),
                );
            }
            val
        }).collect()
    } else {
        items.iter().map(|item| {
            let obj = item.as_object().expect("reference-typed vector requires Object values");
            let raw = obj.as_raw() as usize;
            unsafe { com_to_usize(raw as *mut c_void) }
        }).collect()
    };
    new_vector(packed, is_value_type, iids)
}

/// Create an IVector<T> COM object from a Vec of IUnknown items (reference types).
pub fn create_vector(items: Vec<IUnknown>, iids: VectorIids) -> IUnknown {
    let raw_items: Vec<usize> = items.into_iter().map(|obj| obj.into_raw() as usize).collect();
    new_vector(raw_items, false, iids)
}

/// Create an IVector<T> COM object for value types (structs ≤ pointer size).
pub fn create_value_vector(items: Vec<Vec<u8>>, elem_size: usize, iids: VectorIids) -> IUnknown {
    assert!(
        items.is_empty() || elem_size <= std::mem::size_of::<usize>(),
        "create_value_vector: elem_size {} exceeds pointer size; not yet supported",
        elem_size
    );
    let packed: Vec<usize> = items.iter().map(|bytes| {
        let mut val: usize = 0;
        unsafe {
            std::ptr::copy_nonoverlapping(
                bytes.as_ptr(),
                &mut val as *mut usize as *mut u8,
                bytes.len().min(std::mem::size_of::<usize>()),
            );
        }
        val
    }).collect();
    new_vector(packed, true, iids)
}

fn new_vector(items: Vec<usize>, is_value_type: bool, iids: VectorIids) -> IUnknown {
    let vector = Box::new(SingleThreadedVector {
        vtable_iterable: &SingleThreadedVector::ITERABLE_VTBL,
        vtable_vector: &SingleThreadedVector::VECTOR_VTBL,
        ref_count: windows_core::imp::RefCount::new(1),
        items: RefCell::new(items),
        is_value_type,
        iids,
    });
    unsafe { IUnknown::from_raw(Box::into_raw(vector) as *mut c_void) }
}

// ======================================================================
// Tests
// ======================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::metadata_table::MetadataTable;

    #[test]
    fn test_vector_basic_operations() {
        // Create a vector of IUnknown items using Uri objects
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        let table = MetadataTable::new();
        let iids = table.vector_iids(&table.object());

        // Create Uri objects as test items
        let uri1 = windows::Foundation::Uri::CreateUri(windows_core::h!("https://example.com/1")).unwrap();
        let uri2 = windows::Foundation::Uri::CreateUri(windows_core::h!("https://example.com/2")).unwrap();
        let uri3 = windows::Foundation::Uri::CreateUri(windows_core::h!("https://example.com/3")).unwrap();

        let items: Vec<IUnknown> = vec![
            uri1.cast().unwrap(),
            uri2.cast().unwrap(),
            uri3.cast().unwrap(),
        ];

        let vector = create_vector(items, iids.clone());

        // Test QI for IVector
        let mut vec_ptr = std::ptr::null_mut();
        unsafe { vector.query(&iids.vector, &mut vec_ptr) }.ok().unwrap();
        assert!(!vec_ptr.is_null());

        // Test QI for IIterable
        let mut iter_ptr = std::ptr::null_mut();
        unsafe { vector.query(&iids.iterable, &mut iter_ptr) }.ok().unwrap();
        assert!(!iter_ptr.is_null());

        // Test get_Size via raw vtable call
        let vec_obj = unsafe { IUnknown::from_raw(vec_ptr) };
        let vtbl = unsafe { *(vec_ptr as *const *const VectorVtbl) };
        let mut size: u32 = 0;
        let hr = unsafe { ((*vtbl).get_size)(vec_ptr, &mut size) };
        assert_eq!(hr, S_OK);
        assert_eq!(size, 3);

        // Test get_At
        let mut item_ptr: *mut c_void = std::ptr::null_mut();
        let hr = unsafe { ((*vtbl).get_at)(vec_ptr, 0, &mut item_ptr) };
        assert_eq!(hr, S_OK);
        assert!(!item_ptr.is_null());
        // Release the item
        let _ = unsafe { IUnknown::from_raw(item_ptr) };

        // Test get_At out of bounds
        let hr = unsafe { ((*vtbl).get_at)(vec_ptr, 10, &mut item_ptr) };
        assert_eq!(hr, E_BOUNDS);

        // Release vector interface ref
        drop(vec_obj);
        // Release iterable interface ref
        let _ = unsafe { IUnknown::from_raw(iter_ptr) };
    }

    #[test]
    fn test_vector_iid_computation() {
        let table = MetadataTable::new();

        // IVector<String> IID should match the known PIID computation
        let iids = table.vector_iids(&table.hstring());

        // Verify all IIDs are non-zero (they should be computed from SHA-1)
        assert_ne!(iids.iterable, GUID::zeroed());
        assert_ne!(iids.vector, GUID::zeroed());
        assert_ne!(iids.vector_view, GUID::zeroed());
        assert_ne!(iids.iterator, GUID::zeroed());

        // All should be different from each other
        assert_ne!(iids.iterable, iids.vector);
        assert_ne!(iids.vector, iids.vector_view);
        assert_ne!(iids.vector_view, iids.iterator);
    }

    #[test]
    fn test_vector_append_and_clear() {
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        let table = MetadataTable::new();
        let iids = table.vector_iids(&table.object());

        // Start with empty vector
        let vector = create_vector(Vec::new(), iids.clone());

        // QI to IVector
        let mut vec_ptr = std::ptr::null_mut();
        unsafe { vector.query(&iids.vector, &mut vec_ptr) }.ok().unwrap();
        let vtbl = unsafe { *(vec_ptr as *const *const VectorVtbl) };

        // Size should be 0
        let mut size: u32 = 0;
        unsafe { ((*vtbl).get_size)(vec_ptr, &mut size) };
        assert_eq!(size, 0);

        // Append an item
        let uri = windows::Foundation::Uri::CreateUri(windows_core::h!("https://example.com")).unwrap();
        let unk: IUnknown = uri.cast().unwrap();
        let raw = unk.clone().into_raw();
        unsafe { ((*vtbl).append)(vec_ptr, raw) };

        // Size should now be 1
        unsafe { ((*vtbl).get_size)(vec_ptr, &mut size) };
        assert_eq!(size, 1);

        // Clear
        unsafe { ((*vtbl).clear)(vec_ptr) };
        unsafe { ((*vtbl).get_size)(vec_ptr, &mut size) };
        assert_eq!(size, 0);

        let _ = unsafe { IUnknown::from_raw(vec_ptr) };
    }

    #[test]
    fn test_vector_iterator() {
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        let table = MetadataTable::new();
        let iids = table.vector_iids(&table.object());

        let uri1 = windows::Foundation::Uri::CreateUri(windows_core::h!("https://example.com/1")).unwrap();
        let uri2 = windows::Foundation::Uri::CreateUri(windows_core::h!("https://example.com/2")).unwrap();

        let items: Vec<IUnknown> = vec![
            uri1.cast().unwrap(),
            uri2.cast().unwrap(),
        ];

        let vector = create_vector(items, iids.clone());

        // QI to IIterable
        let mut iter_iface_ptr = std::ptr::null_mut();
        unsafe { vector.query(&iids.iterable, &mut iter_iface_ptr) }.ok().unwrap();
        let iterable_vtbl = unsafe { *(iter_iface_ptr as *const *const IterableVtbl) };

        // Call First()
        let mut iterator_ptr: *mut c_void = std::ptr::null_mut();
        unsafe { ((*iterable_vtbl).first)(iter_iface_ptr, &mut iterator_ptr) };
        assert!(!iterator_ptr.is_null());

        let iter_vtbl = unsafe { *(iterator_ptr as *const *const IteratorVtbl) };

        // HasCurrent should be true
        let mut has_current = false;
        unsafe { ((*iter_vtbl).get_has_current)(iterator_ptr, &mut has_current) };
        assert!(has_current);

        // MoveNext
        let mut has_next = false;
        unsafe { ((*iter_vtbl).move_next)(iterator_ptr, &mut has_next) };
        assert!(has_next); // second item

        unsafe { ((*iter_vtbl).move_next)(iterator_ptr, &mut has_next) };
        assert!(!has_next); // past end

        let _ = unsafe { IUnknown::from_raw(iterator_ptr) };
        let _ = unsafe { IUnknown::from_raw(iter_iface_ptr) };
    }
}
