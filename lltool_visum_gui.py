import wx
import luftlinientool as llt
from pathlib import Path
import logging


## create a simple windows frame (sometimes called form)
class Frame1(wx.Frame):

    # pos=(ulcX,ulcY) size=(width,height) in pixels
    #

    #
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, -1, title, pos=(150, 150), size=(768, 480))

        # ===== Attribute =====
        self.button_vfs_active = None
        self.llt_calculator = None  # llt.LuftlinienCalculator()
        self.iVFS = range(6)
        self.BezirkeEingelesen = False

        self.default_k_nachbar = 1
        self.default_anz_vf = 0

        # ===== Menü ====
        # == Menü Design ==
        # create a menubar at the top of the user frame
        menu_bar = wx.MenuBar()

        # create a menu ...
        menu = wx.Menu()

        # ... add an item to the menu
        #
        # \tAlt-X creates an accelerator for Exit (Alt + x keys)
        # the third parameter is an optional hint that shows up in
        # the statusbar when the cursor moves across this menu item

        einlesen = menu.Append(-1, "&Einlesen Bezirke")
        calculate = menu.Append(-1, "&Berechne Luftlinien-Netz")
        menu.AppendSeparator()
        reset_results = menu.Append(-1, "&Matrix initialisieren")
        show_help = menu.Append(-1, "&Info")
        menu.AppendSeparator()
        default = menu.Append(-1, "&Set Default Values")
        menu.AppendSeparator()
        # menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Exit the program")

        # == Menü Funktionalität ==
        # bind the menu event to an event handler, share QuitBtn event
        self.Bind(wx.EVT_MENU, self.event_quit_button, id=wx.ID_EXIT)
        self.Bind(wx.EVT_MENU, self.event_import_data, einlesen)
        self.Bind(wx.EVT_MENU, self.event_calculate, calculate)
        self.Bind(wx.EVT_MENU, self.event_info, show_help)
        self.Bind(wx.EVT_MENU, self.event_reset, reset_results)
        self.Bind(wx.EVT_MENU, self.event_set_default, default)

        # put the menu on the menubar
        menu_bar.Append(menu, "&Auswahl")
        self.SetMenuBar(menu_bar)

        # create a status bar at the bottom of the frame
        self.CreateStatusBar()
        # now create a panel (between menubar and statusbar) ...
        panel = wx.Panel(self)
        panel.Layout()

        # == Text-Fenster ==
        self.multilabel = wx.StaticText(panel, -1, "Message-Log", (420, 10))
        self.multiText = wx.TextCtrl(panel, -1, "", (420, 40), (320, 300), style=wx.TE_MULTILINE)
        self.multiText.SetInsertionPoint(0)

        self.multiText.WriteText("Programm zur Erzeugung von Luftlinienverbindungen\n")
        self.multiText.WriteText("Dipl.-Ing. Gerd Schleupen\n")
        self.multiText.WriteText("Universitaet Stuttgart\n")
        self.multiText.WriteText("Institut fuer Strassen- und Verkehrswesen\n")
        self.multiText.WriteText("Lehrstuhl fuer Verkehrsplanung und Verkehrsleittechnik\n")

        # ==== Eingabefelder ====

        # == Aktivierung VFS ==
        wx.StaticText(panel, -1, "Verbindungsfunktionsstufe", (20, 10))
        self.button_vfs_active = {"VFS 0": wx.CheckBox(panel, -1, "VFS 0", (20, 80), (60, 20)),
                                  "VFS I": wx.CheckBox(panel, -1, "VFS I", (20, 120), (60, 20)),
                                  "VFS II": wx.CheckBox(panel, -1, "VFS II", (20, 160), (60, 20)),
                                  "VFS III": wx.CheckBox(panel, -1, "VFS III", (20, 200), (60, 20)),
                                  "VFS IV": wx.CheckBox(panel, -1, "VFS IV", (20, 240), (60, 20)),
                                  "VFS V": wx.CheckBox(panel, -1, "VFS V", (20, 280), (60, 20))}

        # == Auswahl Austauschfunktion je VFS ==
        max_nachbar = 99
        wx.StaticText(panel, -1, "Austauschfunktion", (180, 10))
        wx.StaticText(panel, -1, "n-naechste Nachbarn", (180, 40))

        self.buttons_value_k_nachbar_vfs = {"VFS 0": wx.SpinCtrl(panel, -1, "", (180, 80), (40, -1)),
                                            "VFS I": wx.SpinCtrl(panel, -1, "", (180, 120), (40, -1)),
                                            "VFS II": wx.SpinCtrl(panel, -1, "", (180, 160), (40, -1)),
                                            "VFS III": wx.SpinCtrl(panel, -1, "", (180, 200), (40, -1)),
                                            "VFS IV": wx.SpinCtrl(panel, -1, "", (180, 240), (40, -1)),
                                            "VFS V": wx.SpinCtrl(panel, -1, "", (180, 280), (40, -1))}
        self.buttons_value_k_nachbar_vfs["VFS 0"].SetRange(0, max_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS 0"].SetValue(self.default_k_nachbar)

        self.buttons_value_k_nachbar_vfs["VFS I"].SetRange(0, max_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS I"].SetValue(self.default_k_nachbar)

        self.buttons_value_k_nachbar_vfs["VFS II"].SetRange(0, max_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS II"].SetValue(self.default_k_nachbar)

        self.buttons_value_k_nachbar_vfs["VFS III"].SetRange(0, max_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS III"].SetValue(self.default_k_nachbar)

        self.buttons_value_k_nachbar_vfs["VFS IV"].SetRange(0, max_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS IV"].SetValue(self.default_k_nachbar)

        self.buttons_value_k_nachbar_vfs["VFS V"].SetRange(0, max_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS V"].SetValue(self.default_k_nachbar)

        max_vf = 99

        wx.StaticText(panel, -1, "Versorgungsfunktion", (300, 10))
        wx.StaticText(panel, -1, "n-naechst Entfernte", (300, 40))

        self.buttons_value_n_versorger = {"VFS 0": wx.SpinCtrl(panel, -1, "", (300, 80), (40, -1)),
                                          "VFS I": wx.SpinCtrl(panel, -1, "", (300, 120), (40, -1)),
                                          "VFS II": wx.SpinCtrl(panel, -1, "", (300, 160), (40, -1)),
                                          "VFS III": wx.SpinCtrl(panel, -1, "", (300, 200), (40, -1)),
                                          "VFS IV": wx.SpinCtrl(panel, -1, "", (300, 240), (40, -1)),
                                          "VFS V": wx.SpinCtrl(panel, -1, "", (300, 280), (40, -1))}

        self.buttons_value_n_versorger["VFS 0"].SetRange(0, max_vf)
        self.buttons_value_n_versorger["VFS 0"].SetValue(self.default_anz_vf)

        self.buttons_value_n_versorger["VFS I"].SetRange(0, max_vf)
        self.buttons_value_n_versorger["VFS I"].SetValue(self.default_anz_vf)

        self.buttons_value_n_versorger["VFS II"].SetRange(0, max_vf)
        self.buttons_value_n_versorger["VFS II"].SetValue(self.default_anz_vf)

        self.buttons_value_n_versorger["VFS III"].SetRange(0, max_vf)
        self.buttons_value_n_versorger["VFS III"].SetValue(self.default_anz_vf)

        self.buttons_value_n_versorger["VFS IV"].SetRange(0, max_vf)
        self.buttons_value_n_versorger["VFS IV"].SetValue(self.default_anz_vf)

        self.buttons_value_n_versorger["VFS V"].SetRange(0, max_vf)
        self.buttons_value_n_versorger["VFS V"].SetValue(self.default_anz_vf)

    def event_set_default(self, event):

        self.buttons_value_k_nachbar_vfs["VFS 0"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS I"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS II"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS III"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS IV"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS V"].SetValue(self.default_k_nachbar)

        self.buttons_value_n_versorger["VFS 0"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS I"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS II"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS III"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS IV"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS V"].SetValue(self.default_anz_vf)

        self.SetStatusText('Default-Werte hergestellt')
        self.WriteLogText('Default-Werte hergestellt')

    def event_calculate(self, event):
        a = 1

    def event_quit_button(self, evet):
        self.Close()

    def event_import_data(self, event):

        # Falls Visum existiert -> nichts
        # ansonsten Fenster öffnen, mit dem Datei ausgewählt werden kann
        try:
            # testet ob die Variable Visum existiert
            global Visum
            Visum
            name = Visum.UserPreferences.DocumentName
            use_visum = True
        except NameError:
            a = 1 # todo
            source =  Path(r"S:\Mitarbeiter\Schilling") / 'ZentraleOrteBW_Bezirke.ver'

            # Erstellen einer Calculator Instanz
            if source.suffix == ".ver":
                use_visum = True
                Visum = llt.open_visum(source)
            else:
                print("nicht implementiert")

        if use_visum:
            self.llt_calculator = llt.LuftlinienCalculator(Visum, attr_quelle="Quelle", attr_ziel="Ziel", anz_versorger=1,
                                            max_entfernung=1)
        else:
            logging.warning("Umgang mit Nichtvisum Dateien ist nicht implementiert")


    def event_info(self, event):
        self.SetStatusText('Info')
        self.WriteLogText('Info')

    def event_reset(self, event):
        if self.llt_calculator is None:
            a = 1  # todo
        else:
            self.llt_calculator.init_results()


app = wx.App(False)
frame = Frame1(None, "Erstellen von Verbindungsfunktionsstufen-Luftliniennetzen ")
frame.Show(True)
app.MainLoop()

#
