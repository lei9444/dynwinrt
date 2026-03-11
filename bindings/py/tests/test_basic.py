import dynwinrt_py


def test_sum_as_string():
    result = dynwinrt_py.sum_as_string(2, 3)
    assert result == "5"


def test_sum_via_interop():
    result = dynwinrt_py.sum_via_interop(2.5, 3.5)
    assert result == 6.0


def test_process_value_integer():
    result = dynwinrt_py.process_value(42)
    assert result == "Integer: 42"


def test_process_value_float():
    result = dynwinrt_py.process_value(3.14)
    assert result == "Float: 3.14"


def test_process_value_string():
    result = dynwinrt_py.process_value("hello")
    assert result == "String: hello"


def test_process_value_list():
    result = dynwinrt_py.process_value([1, 2, 3])
    assert result == "List with 3 items"


def test_process_value_dict():
    result = dynwinrt_py.process_value({"key": "value"})
    assert result == "Dictionary"


def test_check_type_int():
    result = dynwinrt_py.check_type(42)
    assert result == "int"


def test_check_type_float():
    result = dynwinrt_py.check_type(3.14)
    assert result == "float"


def test_check_type_none():
    result = dynwinrt_py.check_type(None)
    assert result == "None"


def test_check_type_other():
    result = dynwinrt_py.check_type("string")
    assert result == "other"
