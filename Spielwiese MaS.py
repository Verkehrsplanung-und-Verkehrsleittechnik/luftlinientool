# from VisumOverlay import *
#
# import logging
#
#
# if __name__ == '__main__':
#     from pathlib import Path
#     import luftlinientool as llt
#
#     # Parameterübergabe
#     path_source = Path(r"C:\Users\ac128405\Desktop\Software\Luftlinientool\Beispielnetz")
#     file_source = 'ZentraleOrteBW_Bezirke.ver'
#
#     # Settings Logging
#     path_logfile = Path(__file__)
#     path_logfile = Path(r"S:\Forschung\BASt_EmV\13_Skripte") / "Logfiles" / path_logfile.name.replace(".py", ".log")
#
#     root_logger = logging.getLogger()
#     root_logger.setLevel(logging.INFO)
#     logger_format = logging.Formatter("%(asctime)s %(levelname)s: %(message)s", datefmt="%d.%m.%Y %I:%M:%S %p")
#     # Output in Konsole & Logfile
#     stream_handler = logging.StreamHandler()
#     stream_handler.setFormatter(logger_format)
#     file_handler = logging.FileHandler(path_logfile, mode="w")
#     file_handler.setFormatter(logger_format)
#     # add handles to logger
#     root_logger.addHandler(file_handler)
#     root_logger.addHandler(stream_handler)
#
#     source = path_source / file_source
#
#     if source.suffix == ".ver":
#         Visum = llt.open_visum(source)
#         ltt1 = llt.LuftlinienCalculator(Visum, attr_quelle="Quelle", attr_ziel="Ziel", anz_versorger=1, max_entfernung=1)
#     else:
#         print("nicht implementiert")
#
#     ltt1.calculate_main()
#     ltt1.export_matrix(visum=ltt1.visum)
#     ltt1.export_net(visum=ltt1.visum, links_additive=False)
#
#     ltt1.delete_unused_nodes()
#
#     del Visum
#
#

import wx

# Define the tab content as classes:
class TabOne(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        t = wx.StaticText(self, -1, "This is the first tab", (20,20))

class TabTwo(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        t = wx.StaticText(self, -1, "This is the second tab", (20,20))

class TabThree(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        t = wx.StaticText(self, -1, "This is the third tab", (20,20))

class TabFour(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        t = wx.StaticText(self, -1, "This is the last tab", (20,20))


class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="wxPython tabs example @pythonspot.com")

        # Create a panel and notebook (tabs holder)
        p = wx.Panel(self)
        nb = wx.Notebook(p)

        # Create the tab windows
        tab1 = TabOne(nb)
        tab2 = TabTwo(nb)
        tab3 = TabThree(nb)
        tab4 = TabFour(nb)

        # Add the windows to tabs and name them.
        nb.AddPage(tab1, "Tab 1")
        nb.AddPage(tab2, "Tab 2")
        nb.AddPage(tab3, "Tab 3")
        nb.AddPage(tab4, "Tab 4")

        # Set noteboook in a sizer to create the layout
        sizer = wx.BoxSizer()
        sizer.Add(nb, 1, wx.EXPAND)
        p.SetSizer(sizer)


if __name__ == "__main__":
    app = wx.App()
    MainFrame().Show()
    app.MainLoop()