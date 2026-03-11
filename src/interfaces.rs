use std::sync::Arc;
use crate::signature::{InterfaceSignature, MethodSignature};
use crate::metadata_table::MetadataTable;

pub fn uri_factory(reg: &Arc<MetadataTable>) -> InterfaceSignature {
    let mut vtable = InterfaceSignature::define_from_iinspectable("", Default::default(), reg);
    vtable.add_method(
        MethodSignature::new(reg)
            .add_in(reg.hstring())
            .add_out(reg.object()),
    );
    vtable
}

pub fn uri_vtable(reg: &Arc<MetadataTable>) -> InterfaceSignature {
    let mut vtable = InterfaceSignature::define_from_iinspectable(
        "Windows.Foundation.IUriRuntimeClass",
        Default::default(),
        reg,
    );
    vtable
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 6 get_AbsoluteUri
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 7 get_DisplayUri
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 8 get_Domain
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 9 get_Extension
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 10 get_Fragment
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 11 get_Host
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 12 get_Password
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 13 get_Path
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 14 get_Query
        .add_method(MethodSignature::new(reg)) // 15 get_QueryParsed
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 16 get_RawUri
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 17 get_SchemeName
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 18 get_UserName
        .add_method(MethodSignature::new(reg).add_out(reg.i32_type())) // 19 get_Port
        .add_method(MethodSignature::new(reg)); // 20 get_Suspicious;
    vtable
}

pub fn IAsyncOperationWithProgress(reg: &Arc<MetadataTable>) -> InterfaceSignature {
    let mut vtable = InterfaceSignature::define_from_iinspectable(
        "Windows.Foundation.IAsyncOperationWithProgress",
        Default::default(),
        reg,
    );
    vtable
        .add_method(MethodSignature::new(reg)) // 6 SetProgress
        .add_method(MethodSignature::new(reg)) // 7 GetProgress
        .add_method(MethodSignature::new(reg)) // 8 SetCompleted
        .add_method(MethodSignature::new(reg)) // 9 GetCompleted
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())); // 10 GetResults
    vtable
}

pub fn IAsyncOperation(reg: &Arc<MetadataTable>) -> InterfaceSignature {
    let mut vtable = InterfaceSignature::define_from_iinspectable(
        "Windows.Foundation.IAsyncOperation",
        Default::default(),
        reg,
    );
    vtable
        .add_method(MethodSignature::new(reg)) // 6 SetCompleted
        .add_method(MethodSignature::new(reg)) // 7 GetCompleted
        .add_method(MethodSignature::new(reg).add_out(reg.object())); // 8 GetResults
    vtable
}

pub fn FileOpenPickerFactory(reg: &Arc<MetadataTable>) -> InterfaceSignature {
    let mut vtable = InterfaceSignature::define_from_iinspectable(
        "Windows.Storage.Pickers.IFileOpenPickerFactory",
        Default::default(),
        reg,
    );
    vtable.add_method(
        MethodSignature::new(reg)
            .add_in(reg.i64_type())
            .add_out(reg.object()),
    ); // 6 CreateWithMode
    vtable
}

pub fn PickFileResult(reg: &Arc<MetadataTable>) -> InterfaceSignature {
    let mut vtable = InterfaceSignature::define_from_iinspectable(
        "Windows.Storage.Pickers.PickFileResult",
        Default::default(),
        reg,
    );
    vtable.add_method(MethodSignature::new(reg).add_out(reg.hstring())); // 6 get_File
    vtable
}

pub fn FileOpenPicker(reg: &Arc<MetadataTable>) -> InterfaceSignature {
    let mut vtable = InterfaceSignature::define_from_iinspectable(
        "Windows.Storage.Pickers.IFileOpenPicker",
        Default::default(),
        reg,
    );
    vtable
        .add_method(MethodSignature::new(reg).add_in(reg.i32_type())) // 6 put_ViewMode
        .add_method(MethodSignature::new(reg).add_out(reg.i32_type())) // 7 get_ViewMode
        .add_method(MethodSignature::new(reg).add_in(reg.object())) // 8 put_SuggestedStartLocation
        .add_method(MethodSignature::new(reg).add_out(reg.object())) // 9 get_SuggestedStartLocation
        .add_method(MethodSignature::new(reg).add_in(reg.hstring())) // 10 put_CommitButtonText
        .add_method(MethodSignature::new(reg).add_out(reg.hstring())) // 11 get_CommitButtonText
        .add_method(MethodSignature::new(reg).add_out(reg.object())) // 12 get_FileTypeFilter
        .add_method(MethodSignature::new(reg).add_out(reg.object())); // 13 PickSingleFileAsync
    vtable
}
