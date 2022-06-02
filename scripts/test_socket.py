from pytemscript.utils.gatan_socket import GatanSocket


def test1():
    g = GatanSocket()
    ver = g.GetDMVersion()
    print('Version', ver)

    s = 'Result("Hello world\\n")'
    g.ExecuteScript(s)


if __name__ == '__main__':
    test1()
