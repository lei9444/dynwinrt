import dynwinrt_py
from dynwinrt_py import (
    DynWinRTType,
    DynWinRTMethodSig,
    DynWinRTValue,
    WinGUID,
    DynWinRTArray,
    DynWinRTStruct,
    ro_initialize,
)


def test_ro_initialize():
    """RoInitialize should succeed (or already initialized)."""
    ro_initialize(1)


def test_primitive_types():
    """All primitive type factories should return DynWinRTType instances."""
    assert DynWinRTType.i32_type() is not None
    assert DynWinRTType.i64_type() is not None
    assert DynWinRTType.hstring() is not None
    assert DynWinRTType.object() is not None
    assert DynWinRTType.f32_type() is not None
    assert DynWinRTType.f64_type() is not None
    assert DynWinRTType.u8_type() is not None
    assert DynWinRTType.u16_type() is not None
    assert DynWinRTType.u32_type() is not None
    assert DynWinRTType.u64_type() is not None
    assert DynWinRTType.i8_type() is not None
    assert DynWinRTType.i16_type() is not None
    assert DynWinRTType.bool_type() is not None


def test_guid_parse():
    """WinGUID.parse should parse valid GUIDs."""
    guid = WinGUID.parse("9e365e57-48b2-4160-956f-c7385120bbfc")
    assert guid is not None
    assert "WinGUID" in repr(guid)


def test_value_from_hstring():
    v = DynWinRTValue.from_hstring("hello")
    assert str(v) == "hello"
    assert v.to_string() == "hello"


def test_value_from_i32():
    v = DynWinRTValue.from_i32(42)
    assert v.to_int() == 42
    assert v.to_string() == "42"


def test_value_from_i64():
    v = DynWinRTValue.from_i64(123456789)
    assert v.to_int() == 123456789


def test_value_from_f64():
    v = DynWinRTValue.from_f64(3.14)
    assert abs(v.to_float() - 3.14) < 1e-10


def test_value_from_bool():
    v = DynWinRTValue.from_bool(True)
    assert v.to_int() == 1


def test_method_sig_builder():
    """MethodSig builder chain should work."""
    sig = DynWinRTMethodSig()
    sig2 = sig.add_in(DynWinRTType.hstring())
    sig3 = sig2.add_out(DynWinRTType.object())
    assert sig3 is not None


def test_register_interface_and_add_method():
    """Register an interface and add a method."""
    iid = WinGUID.parse("9e365e57-48b2-4160-956f-c7385120bbfc")
    iface = DynWinRTType.register_interface("TestInterface", iid)
    sig = DynWinRTMethodSig().add_out(DynWinRTType.hstring())
    iface2 = iface.add_method("GetName", sig)
    handle = iface2.method(6)
    assert handle is not None


def test_uri_dynamic_invocation():
    """Full round-trip: create Uri via activation factory, read properties."""
    ro_initialize(1)

    # IUriRuntimeClassFactory IID
    factory_iid = WinGUID.parse("44a9796f-723e-4fdf-a218-033e75b0c084")
    # IUriRuntimeClass IID
    uri_iid = WinGUID.parse("9e365e57-48b2-4160-956f-c7385120bbfc")

    # Register IUriRuntimeClassFactory interface
    factory_type = DynWinRTType.register_interface("IUriRuntimeClassFactory", factory_iid)
    create_uri_sig = DynWinRTMethodSig().add_in(DynWinRTType.hstring()).add_out(DynWinRTType.object())
    factory_type = factory_type.add_method("CreateUri", create_uri_sig)

    # Register IUriRuntimeClass interface
    uri_type = DynWinRTType.register_interface("IUriRuntimeClass", uri_iid)
    uri_type = uri_type.add_method("get_AbsoluteUri", DynWinRTMethodSig().add_out(DynWinRTType.hstring()))
    uri_type = uri_type.add_method("get_DisplayUri", DynWinRTMethodSig().add_out(DynWinRTType.hstring()))

    # Get activation factory
    factory = DynWinRTValue.activation_factory("Windows.Foundation.Uri")

    # Cast to IUriRuntimeClassFactory
    factory_obj = factory.cast(factory_iid)

    # Create a Uri
    create_method = factory_type.method_by_name("CreateUri")
    uri_obj = create_method.invoke(factory_obj, [DynWinRTValue.from_hstring("https://example.com/path")])

    # Cast to IUriRuntimeClass and read AbsoluteUri
    uri_casted = uri_obj.cast(uri_iid)
    get_abs = uri_type.method_by_name("get_AbsoluteUri")
    abs_uri = get_abs.invoke(uri_casted, [])
    assert abs_uri.to_string() == "https://example.com/path"


def test_array_from_i32():
    arr = DynWinRTArray.from_i32_values([1, 2, 3, 4, 5])
    assert len(arr) == 5
    assert arr.get(0).to_int() == 1
    assert arr.get(4).to_int() == 5
    assert arr.to_i32_list() == [1, 2, 3, 4, 5]


def test_array_from_f64():
    arr = DynWinRTArray.from_f64_values([1.5, 2.5, 3.5])
    assert len(arr) == 3
    assert arr.to_f64_list() == [1.5, 2.5, 3.5]


def test_array_from_u8():
    arr = DynWinRTArray.from_u8_values([0, 127, 255])
    assert len(arr) == 3
    assert arr.to_u8_list() == bytes([0, 127, 255])


def test_array_to_value():
    """Array can be wrapped as DynWinRTValue."""
    arr = DynWinRTArray.from_i32_values([10, 20])
    val = arr.to_value()
    assert val.is_array()
    roundtrip = val.as_array()
    assert len(roundtrip) == 2


def test_struct_create_and_field_access():
    """Create a struct and get/set fields."""
    typ = DynWinRTType.struct_type([DynWinRTType.i32_type(), DynWinRTType.f64_type()])
    s = DynWinRTStruct.create(typ)
    assert s.get_i32(0) == 0
    s.set_i32(0, 42)
    assert s.get_i32(0) == 42
    s.set_f64(1, 3.14)
    assert abs(s.get_f64(1) - 3.14) < 1e-10


def test_struct_to_value():
    typ = DynWinRTType.struct_type([DynWinRTType.u32_type()])
    s = DynWinRTStruct.create(typ)
    s.set_u32(0, 99)
    val = s.to_value()
    assert val.is_struct()
