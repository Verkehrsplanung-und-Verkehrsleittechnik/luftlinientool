import wx
import luftlinientool as llt
from pathlib import Path
import logging

# Erstellt das Layout für das GUI

# ===== Hilfsfkt =====

## lädt alle Visumattribute
def get_attr_zones(Visum):
    list_attr = Visum.Net.Zones.Attributes.GetAll
    list_attr_names = [attr.Code for attr in list_attr]

    return list_attr_names

# Definiert das komplette Fenster, erzeugt die einzelnen Bestandteile und verbindet diese mit der Logik
class LLTFrame(wx.Frame):
    def __init__(self):
        super().__init__(parent=None)

        # ===== Attribute =====
        self.buttons_value_n_versorger = None
        self.cb_quelle = None
        self.cb_ziel = None
        self.cb_vfs = None
        self.buttons_vfs_value = None
        self.buttons_value_k_nachbar_vfs = None
        self.button_vfs_active = None
        self.llt_calculator = None  # llt.LuftlinienCalculator()
        self.default_k_nachbar = 1
        self.default_anz_vf = 0
        self.attr_vfs = "TypeNo"
        self.attr_quelle = None
        self.attr_ziel = None

        # Falls Visum existiert -> nichts
        # ansonsten Fenster öffnen, mit dem Datei ausgewählt werden kann
        try:
            # testet ob die Variable Visum existiert
            global Visum
            Visum
        except NameError:
            import win32com.client as com
            defDir = Path.cwd()
            with wx.FileDialog(self,
                                'Wähle ein Visumnetz',
                                str(defDir),
                                wildcard='Versiondateien (*.ver)|*.ver',
                                style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dlg:
                if file_dlg.ShowModal() == wx.ID_CANCEL:
                    return  # the user changed their mind
                source =  Path(file_dlg.Path)

            if source.suffix == ".ver":
                version = 220
                Visum = com.Dispatch(f"Visum.Visum.{version}")
                logging.info('open visum file: {}'.format(source))
                Visum.LoadVersion(source)

        self.visum = Visum
        self.list_attr = get_attr_zones(self.visum)

        self.attr_quelle = None
        self.attr_ziel = None
        self.attr_vfs = "TypeNo"

        self.__set_layout__()
        self.__set_properties__()

        self.Show()

    def __set_properties__(self):
        self.SetTitle("Erstellen von Verbindungsfunktionsstufen-Luftliniennetzen")
        self.SetMinSize((1000,500))

        max_value_vfs = int(self.visum.Net.AttValue(f"Max:Zones\{self.attr_vfs}"))
        idx = 0
        for btn in self.buttons_vfs_value.values():
            btn.SetRange(0, max_value_vfs)
            btn.SetValue(idx)
            idx += 1

        self.event_set_default()

    def __set_layout__(self):
        # Layout
        # oben Menübar/Toolbar
        # dann

        # create a panel (between menubar and statusbar) ...
        self.panel = wx.Panel(self)
        self.notebook = wx.Notebook(self.panel)

        # create tabs
        self.tabMain = MainTab(self.notebook)
        self.tabLog = LogTab(self.notebook)

        self.notebook.AddPage(self.tabMain, "Main")
        self.notebook.AddPage(self.tabLog, "Log")

        # create a menubar at the top of the user frame
        self.menu_bar = wx.MenuBar()

        # create a menu ...
        self.menu = wx.Menu()
        self.menu.Append(11, "&Daten einlesen")
        self.menu.Append(12, "&Berechnung Luftlinien-Netz")
        self.menu.AppendSeparator()
        self.menu.Append(13, "&Ergebnisse initialisieren")
        self.menu.Append(14, "&Defaultwerte übernehmen")
        self.menu.AppendSeparator()
        self.menu.Append(15, "&Info")

        self.menu.AppendSeparator()
        # put the menu on the menubar
        self.menu_bar.Append(self.menu, "&Auswahl")
        self.SetMenuBar(self.menu_bar)

        self.toolbar = self.CreateToolBar(style=wx.TB_TEXT | wx.TB_NOICONS)

        # Workaroun keine Bilder zur Verfügung: Leeres Bitmap Objekt
        self.toolbar.AddTool(101, 'Daten einlesen', wx.Bitmap())
        self.toolbar.AddTool(102, 'Berechnung Luftlinien-Netz', wx.Bitmap())
        self.toolbar.AddTool(103, 'Ergebnisse initialisieren', wx.Bitmap())
        self.toolbar.AddTool(104, 'Defaultwerte', wx.Bitmap())
        self.toolbar.AddTool(105, 'Info', wx.Bitmap())
        self.toolbar.Realize()


        # # # create tool bar
        # # toolbar = self.CreateToolBar()
        # # qtool = toolbar.AddTool(wx.ID_ANY, 'Quit', wx.Bitmap('Exit.bmp'))
        # # toolbar.Realize()

        # create a status bar at the bottom of the frame
        self.CreateStatusBar()

        # Set noteboook in a sizer to create the layout
        sizer = wx.BoxSizer()
        sizer.Add(self.notebook, 1, wx.EXPAND)
        self.panel.SetSizer(sizer)

    def __bind_events__(self):
        # Event Handler
        # bind the menu event to an event handler, share QuitBtn event
        self.Bind(wx.EVT_CLOSE, self.event_quit_button)
        self.Bind(wx.EVT_MENU, self.event_import_data, 11)
        self.Bind(wx.EVT_MENU, self.event_calculate, 12)
        self.Bind(wx.EVT_MENU, self.event_info, 15)
        self.Bind(wx.EVT_MENU, self.event_reset, 13)
        self.Bind(wx.EVT_MENU, self.event_set_default, 14)

        self.toolbar.Bind(wx.EVT_TOOL, self.event_import_data, id=101)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_calculate, id=102)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_reset, id=103)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_set_default, id=104)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_info, id=105)

    def event_set_default(self, event=None):
        # funktionsfähig, ggf Default Attributwerte VFS ergänzen

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

        self.cb_vfs.SetValue("TypeNo")
        self.cb_quelle.SetValue('')
        self.cb_ziel.SetValue('')

        self.SetStatusText('Default-Werte hergestellt')

    def event_choose_attr(self, event):
        attr = event.GetEventObject().GetStringSelection()

        if event.GetEventObject().Label == 'attr_vfs':
            self.attr_vfs = attr
        elif event.GetEventObject().Label == 'attr_quelle':
            if attr == '':
                self.attr_quelle = None
            else:
                self.attr_quelle = attr
        elif event.GetEventObject().Label == 'attr_ziel':
            if attr == '':
                self.attr_ziel = None
            else:
                self.attr_ziel = attr
        else:
            logging.warning("sollte nie passieren")
            
        self.SetStatusText("Attribut übernommen, Bezirke neu importieren nicht vergessen")

    def event_calculate(self, event):
        # Vorgehen
        # 1. Update der vorgegebenen parameter, falls was geändert wurde
        # 2. berechnen
        self.update_param_vfs()
        self.llt_calculator.calculate_main()

        # Statusleiste
        self.SetStatusText('Berechnung durchgeführt')

    def event_quit_button(self, event):
        # del self.visum
        #self.stop = True
        self.panel.Destroy()
        self.Destroy()
        wx.GetApp().ExitMainLoop()


    def event_import_data(self, event):
        # funktioniert soweit,
        # Erstellen einer Calculator Instanz
        if self.visum is not None:
            # Init Calculator Instanz
            self.llt_calculator = llt.LuftlinienCalculator(self.visum,
                                                           attr_vfs= self.attr_vfs,
                                                           attr_quelle=self.attr_quelle,
                                                           attr_ziel=self.attr_ziel,
                                                           anz_versorger=1,
                                                           max_entfernung=1)
            # Übergebe aktuelle Parameter
            self.update_param_vfs()
        else:
            logging.warning("Umgang mit Nichtvisum Dateien ist nicht implementiert")
        self.SetStatusText('Daten importiert')

    def event_info(self, event):
        a=1

    def event_reset(self, event):
        if self.llt_calculator is None:
            a = 1  # todo
        else:
            self.llt_calculator.init_results()

        self.SetStatusText('Ergebnisse gelöscht, ggf Ergebnisse neu nach Visum importieren')

    def event_export_results(self, event):
        if self.llt_calculator is not None:
            self.llt_calculator.export_net(visum=self.visum,
                                           links_additive=False)
            self.llt_calculator.export_matrix(visum=self.visum)

    def event_export_net(self, event):
        vfs = event.GetEventObject().vfs

        if self.llt_calculator is not None:
            self.llt_calculator.export_net(list_vfs=[vfs],
                                           visum=self.visum,
                                           links_additive=False)
            self.llt_calculator.delete_unused_nodes()

        self.SetStatusText(f'{vfs}: Net-Datei exportiert und in Visum importiert')


    def event_export_mtx(self, event):
        vfs = event.GetEventObject().vfs

        if self.llt_calculator is not None:
            self.llt_calculator.export_matrix(list_vfs=[vfs],   visum=self.visum)

        self.SetStatusText(f'{vfs}: Matrix in Visum geladen')

    def event_export_master(self, event):
        self.llt_calculator.export_matrix(visum=self.visum)
        self.llt_calculator.export_net(visum=self.visum,
                                       links_additive=False)
        self.llt_calculator.delete_unused_nodes()

        self.SetStatusText(f'die kombinierten Ergebnisse wurden in Visum importiert')

    def update_param_vfs(self):
        if self.llt_calculator is not None:
            list_vfs = [vfs[0] for vfs in self.button_vfs_active.items() if vfs[1].Value > 0]
            dict_anz_versorger = {vfs: self.buttons_value_n_versorger[vfs].Value for vfs in list_vfs}
            dict_max_nachbar = {vfs: self.buttons_value_k_nachbar_vfs[vfs].Value for vfs in list_vfs}
            dict_vfs = {vfs: self.buttons_vfs_value[vfs].Value for vfs in list_vfs}

            self.llt_calculator.nachbarschaftsgrad_vfs = dict_max_nachbar
            self.llt_calculator.anz_versorger_vfs = dict_anz_versorger
            self.llt_calculator.vfs = dict_vfs

class MainTab(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        self.__set_layout__()
        self.__bind_events__()


    def __set_layout__(self):
        # Zeilen mit einzelnen Elementen (vbox_outer)
        # Zeile 1: Bezirksattributauswahl
        # Zeile 2: GridbagSizer mit allem auser Log
        # unten Statusbar

        vbox_outer = wx.BoxSizer(wx.VERTICAL)
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        gridbagsizer1 = wx.GridBagSizer(vgap=10, hgap=50)

        # Auswahl Bezirksattribute
        self.cb_vfs = wx.ComboBox(self, size=(200, -1), choices=self.TopLevelParent.list_attr, style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)
        self.cb_vfs.Label= 'attr_vfs'
        self.TopLevelParent.cb_vfs = self.cb_vfs

        self.cb_quelle = wx.ComboBox(self, size=(200, -1), choices=self.TopLevelParent.list_attr, style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)
        self.cb_quelle.Label= 'attr_quelle'
        self.TopLevelParent.cb_quelle = self.cb_quelle

        self.cb_ziel = wx.ComboBox(self, size=(200, -1), choices=self.TopLevelParent.list_attr, style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)
        self.cb_ziel.Label= 'attr_ziel'
        self.TopLevelParent.cb_ziel = self.cb_ziel

        hbox1.Add(wx.StaticText(self, -1, "Bezirksattribut \n VFS"), 0, wx.ALL|wx.EXPAND,5)
        hbox1.Add(self.cb_vfs, 0, wx.ALL|wx.EXPAND,15)
        hbox1.Add(wx.StaticText(self, -1, "Bezirksattribut \n 'ist Quelle'"), 0, wx.ALL|wx.EXPAND,5)
        hbox1.Add(self.cb_quelle, 0, wx.ALL|wx.EXPAND,15)
        hbox1.Add(wx.StaticText(self, -1, "Bezirksattribut \n 'ist Ziel'"), 0, wx.ALL|wx.EXPAND,5)
        hbox1.Add(self.cb_ziel, 0, wx.ALL|wx.EXPAND,15 )

        # Überschrift Spalte 1
        gridbagsizer1.Add(wx.StaticText(self, -1, "Verbindungsfunktionsstufe"),
                          pos=(0,0), flag=wx.TOP|wx.LEFT|wx.BOTTOM, border=5)
        self.button_vfs_active = {"VFS 0": wx.CheckBox(self, -1, "VFS 0"),
                                  "VFS I": wx.CheckBox(self, -1, "VFS I"),
                                  "VFS II": wx.CheckBox(self, -1, "VFS II"),
                                  "VFS III": wx.CheckBox(self, -1, "VFS III"),
                                  "VFS IV": wx.CheckBox(self, -1, "VFS IV"),
                                  "VFS V": wx.CheckBox(self, -1, "VFS V")}

        tmp_iterator = 1
        for btn in self.button_vfs_active.values():
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 0), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.button_vfs_active = self.button_vfs_active

        # Spalte 2 Angabe Wert je VFS
        gridbagsizer1.Add(wx.StaticText(self, -1, "Attributwert VFS"),
                          pos=(0, 1), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_vfs_value = {"VFS 0": wx.SpinCtrl(self, -1, ""),
                                 "VFS I": wx.SpinCtrl(self, -1, ""),
                                 "VFS II": wx.SpinCtrl(self, -1, ""),
                                 "VFS III": wx.SpinCtrl(self, -1, ""),
                                 "VFS IV": wx.SpinCtrl(self, -1, ""),
                                 "VFS V": wx.SpinCtrl(self, -1, "")}
        tmp_iterator = 1
        for btn in self.buttons_vfs_value.values():
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 1), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_vfs_value = self.buttons_vfs_value

        # Spalte 2 Auswahl Austauschfunktion je VFS
        gridbagsizer1.Add(wx.StaticText(self, -1, "Austauschfunktion \n n-naechste Nachbarn"),
                          pos=(0, 2), flag=wx.ALIGN_CENTER | wx.ALL)

        self.buttons_value_k_nachbar_vfs = {"VFS 0": wx.SpinCtrl(self, -1, ""),
                                            "VFS I": wx.SpinCtrl(self, -1, ""),
                                            "VFS II": wx.SpinCtrl(self, -1, ""),
                                            "VFS III": wx.SpinCtrl(self, -1, ""),
                                            "VFS IV": wx.SpinCtrl(self, -1, ""),
                                            "VFS V": wx.SpinCtrl(self, -1, "")}
        tmp_iterator = 1
        for btn in self.buttons_value_k_nachbar_vfs.values():
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 2), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_value_k_nachbar_vfs = self.buttons_value_k_nachbar_vfs

        # Spalte 3 Versorgunsgfunktion
        gridbagsizer1.Add(
            wx.StaticText(self, -1, "Versorgungsfunktion \n n Versorgungszentren"),
            pos=(0, 3), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_value_n_versorger = {"VFS 0": wx.SpinCtrl(self, -1, ""),
                                          "VFS I": wx.SpinCtrl(self, -1, ""),
                                          "VFS II": wx.SpinCtrl(self, -1, ""),
                                          "VFS III": wx.SpinCtrl(self, -1, ""),
                                          "VFS IV": wx.SpinCtrl(self, -1, ""),
                                          "VFS V": wx.SpinCtrl(self, -1, "")}

        tmp_iterator = 1
        for btn in self.buttons_value_n_versorger.values():
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 3), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_value_n_versorger = self.buttons_value_n_versorger

        # Buttons Export Matrix
        gridbagsizer1.Add(
            wx.StaticText(self, -1, "anlegen in Visum als"),
            pos=(0, 4), span=(1,2), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_export_mat = {"VFS 0": wx.Button(self, -1, "MTX"),
                                   "VFS I": wx.Button(self, -1, "MTX"),
                                   "VFS II": wx.Button(self, -1, "MTX"),
                                   "VFS III": wx.Button(self, -1, "MTX"),
                                   "VFS IV": wx.Button(self, -1, "MTX"),
                                   "VFS V": wx.Button(self, -1, "MTX")}

        tmp_iterator = 1
        for vfs, btn in self.buttons_export_mat.items():
            btn.vfs = vfs
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 4), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        # Buttons Export Net
        self.buttons_export_net = {"VFS 0": wx.Button(self, -1, "Net"),
                                   "VFS I": wx.Button(self, -1, "Net"),
                                   "VFS II": wx.Button(self, -1, "Net"),
                                   "VFS III": wx.Button(self, -1, "Net"),
                                   "VFS IV": wx.Button(self, -1, "Net"),
                                   "VFS V": wx.Button(self, -1, "Net")}
        tmp_iterator = 1
        for vfs, btn in self.buttons_export_net.items():
            btn.vfs = vfs
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 5), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        # Buttons export all
        self.btn_export_master = wx.Button(self, -1, "Import nach Visum alle VFS \n Strecken + Mtx ")
        self.btn_export_master.vfs = 'alle'
        gridbagsizer1.Add(self.btn_export_master,
                          pos=(7,4), span=(3,2), flag= wx.EXPAND)
        # Aufbau Layout
        vbox_outer.Add(hbox1, 0 , wx.ALL | wx.EXPAND, 1)
        vbox_outer.Add(gridbagsizer1,  1, wx.ALL | wx.EXPAND, 6)
        self.SetSizer(vbox_outer)

        # ==== Event binding
    def __bind_events__(self):

        for vfs, btn in self.buttons_export_net.items():
            btn.Bind(wx.EVT_BUTTON, self.TopLevelParent.event_export_net)

        for vfs, btn in self.buttons_export_mat.items():
            btn.Bind(wx.EVT_BUTTON, self.TopLevelParent.event_export_mtx)

        self.btn_export_master.Bind(wx.EVT_BUTTON, self.TopLevelParent.event_export_master)

        self.cb_vfs.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)
        self.cb_quelle.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)
        self.cb_ziel.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)


class LogTab(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(wx.StaticText(self, -1, "Message-Log"), 0, wx.ALL|wx.CENTER, 5)
        # self.multiText = wx.TextCtrl(panel, -1, "", style=wx.TE_MULTILINE)
        # self.multiText.SetInsertionPoint(0)
        #
        # vbox2.Add(self.multiText, 2, wx.ALIGN_CENTER | wx.ALL, 20)
        #

        # ==== Logging =====
        path_logfile = Path(__file__)
        path_logfile = path_logfile.name.replace(".py", ".log")
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)
        logger_format = logging.Formatter("%(asctime)s %(levelname)s: %(message)s", datefmt="%d.%m.%Y %I:%M:%S %p")

        # Output in Konsole & Logfile
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(logger_format)
        file_handler = logging.FileHandler(path_logfile, mode="w")
        file_handler.setFormatter(logger_format)
        # add handles to logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(stream_handler)

        self.log = wx.TextCtrl(self, wx.ID_ANY, size=(700,200),
                          style = wx.TE_MULTILINE|wx.TE_READONLY|wx.HSCROLL |wx.EXPAND)
        self.handler = WxTextCtrlHandler(self.log)
        self.handler.setFormatter(logger_format)
        self.logger.addHandler(self.handler)

        vbox.Add(self.log, 1, wx.ALL | wx.EXPAND, 5)
        self.SetSizer(vbox)


class WxTextCtrlHandler(logging.Handler):
    def __init__(self, ctrl):
        logging.Handler.__init__(self)
        self.ctrl = ctrl

    def emit(self, record):
        s = self.format(record) + '\n'
        wx.CallAfter(self.ctrl.WriteText, s)

if __name__ == '__main__':
    app = wx.App()
    frame = LLTFrame()
    app.MainLoop()
