import dynwinrt_py

def test_winrt_method():
    m = dynwinrt_py.WinRTMethod()
    print(m)

def test_winrt_interface():
    iface = dynwinrt_py.WinRTInterface()
    m = dynwinrt_py.WinRTMethod()
    result = iface.add_method(m)
    print(result)
    
