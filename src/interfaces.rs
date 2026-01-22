use crate::signature::{MethodSignature, VTableSignature};
use crate::types::WinRTType;

pub fn uri_vtable() -> VTableSignature {
    let mut vtable = VTableSignature::new();
    vtable
        .add_method(MethodSignature::new()) // 0 QueryInterface
        .add_method(MethodSignature::new()) // 1 AddRef
        .add_method(MethodSignature::new()) // 2 Release
        .add_method(MethodSignature::new()) // 3 GetIids
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 4 GetRuntimeClassName
        .add_method(MethodSignature::new()) // 5 GetTrustLevel
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 6 get_AbsoluteUri
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 7 get_DisplayUri
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 8 get_Domain
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 9 get_Extension
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 10 get_Fragment
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 11 get_Host
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 12 get_Password
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 13 get_Path
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 14 get_Query
        .add_method(MethodSignature::new()) // 15 get_QueryParsed
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 16 get_RawUri
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 17 get_SchemeName
        .add_method(MethodSignature::new().add_out(WinRTType::HString)) // 18 get_UserName
        .add_method(MethodSignature::new().add_out(WinRTType::I32)) // 19 get_Port
        .add_method(MethodSignature::new()); // 20 get_Suspicious;
    vtable
}
