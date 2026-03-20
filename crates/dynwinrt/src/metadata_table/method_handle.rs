use std::sync::Arc;

use crate::value::WinRTValue;

use super::MetadataTable;

/// A handle to a pre-built method in the MetadataTable's methods arena.
#[derive(Clone)]
pub struct MethodHandle {
    pub(crate) table: Arc<MetadataTable>,
    pub(crate) index: u32,
}

impl std::fmt::Debug for MethodHandle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("MethodHandle")
            .field("index", &self.index)
            .finish()
    }
}

impl MethodHandle {
    pub(crate) fn new(table: Arc<MetadataTable>, index: u32) -> Self {
        MethodHandle { table, index }
    }

    /// Invoke this method on the given COM object with the provided arguments.
    pub fn invoke(
        &self,
        obj: *mut std::ffi::c_void,
        args: &[WinRTValue],
    ) -> crate::result::Result<Vec<WinRTValue>> {
        self.table
            .invoke_method(self.index, obj, args)
            .map_err(crate::result::Error::WindowsError)
    }
}
