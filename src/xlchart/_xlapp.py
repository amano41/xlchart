from win32com.client import DispatchEx, GetObject, gencache, makepy


def _init_excel():
    try:
        gencache.EnsureModule("{00020813-0000-0000-C000-000000000046}", 0, 1, 9)
        return
    except Exception:
        pass

    for spec in (
        "Microsoft Excel 16.0 Object Library",
        "Microsoft Excel 15.0 Object Library",
        "Excel.Application",
    ):
        try:
            makepy.GenerateFromTypeLibSpec(spec)
            return
        except Exception:
            continue

    exist = False
    try:
        if GetObject(Class="Excel.Application"):
            exist = True
    except Exception:
        pass

    xl = None
    try:
        xl = gencache.EnsureDispatch("Excel.Application")
        if not exist:
            xl.Quit()
    finally:
        if xl is not None:
            del xl


def _new_excel(vidible: bool = False):
    xl = DispatchEx("Excel.Application")
    xl.Visible = vidible
    xl.DisplayAlerts = False
    try:
        xl.ScreenUpdating = False
    except Exception:
        pass
    return xl


def _quit_excel(xl):
    if xl is not None:
        xl.DisplayAlerts = True
        try:
            xl.ScreenUpdating = True
        except Exception:
            pass
        xl.Quit()
        del xl
