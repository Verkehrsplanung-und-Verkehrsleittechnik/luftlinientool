## @package llt_GUI.py
# @brief Verwaltet und definiert die grafische Benutzeroberfläche für das Luftlinientool

import wx
import luftlinientool as llt
from pathlib import Path
import logging

# ===== Hilfsfkt =====

## lädt alle Visumattribute
def get_attr_zones(Visum):
    list_attr = Visum.Net.Zones.Attributes.GetAll
    list_attr_id = [attr.ID for attr in list_attr]

    return list_attr_id

# ======= Klassen ======

## Definiert das komplette Fenster, erzeugt die einzelnen Bestandteile und verbindet diese mit der Logik
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
                version = 240
                Visum = com.Dispatch(f"Visum.Visum.{version}")
                logging.info('open visum file: {}'.format(source))
                Visum.LoadVersion(source)

        self.visum = Visum
        self.list_attr = get_attr_zones(self.visum)

        self.attr_quelle = None
        self.attr_ziel = None
        self.attr_vfs = "TypeNo"
        self.attr_dist_fcn = "euclidean"

        self.__set_layout__()
        self.__set_properties__()
        self.__bind_events__()

        self.Show()

    def __set_properties__(self):
        self.SetTitle("Erstellen von Verbindungsfunktionsstufen-Luftliniennetzen")
        self.SetMinSize((1200,500))

        self.__set_values_vfs_buttons__()

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

        # Workaround keine Bilder zur Verfügung: Leeres Bitmap Objekt
        self.toolbar.AddTool(101, 'Daten einlesen', wx.Bitmap())
        self.toolbar.AddTool(102, 'Berechnung Luftlinien-Netz', wx.Bitmap())
        self.toolbar.AddTool(103, 'Ergebnisse initialisieren', wx.Bitmap())
        self.toolbar.AddTool(104, 'Defaultwerte', wx.Bitmap())
        self.toolbar.AddTool(105, 'Info', wx.Bitmap())
        self.toolbar.AddTool(106, 'Filter: eingefügte Strecken', wx.Bitmap())
        self.toolbar.AddTool(107, 'Löschen: eingefügte Strecken', wx.Bitmap())
        self.toolbar.Realize()


        # # # create toolbar
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
        self.Bind(wx.EVT_MENU, self.event_import_data, id=11)
        self.Bind(wx.EVT_MENU, self.event_calculate, id=12)
        self.Bind(wx.EVT_MENU, self.event_info, id=15)
        self.Bind(wx.EVT_MENU, self.event_reset, id=13)
        self.Bind(wx.EVT_MENU, self.event_set_default, id=14)

        self.toolbar.Bind(wx.EVT_TOOL, self.event_import_data, id=101)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_calculate, id=102)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_reset, id=103)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_set_default, id=104)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_info, id=105)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_filter, id=106)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_delete_links, id=107)

    def __set_values_vfs_buttons__(self):
        max_value_vfs = int(self.visum.Net.AttValue(f"Max:Zones\{self.attr_vfs}"))
        idx = 0
        for btn in self.buttons_vfs_value.values():
            btn.SetRange(0, max_value_vfs)
            btn.SetValue(idx)
            idx += 1


    def event_set_default(self, event=None):
        # funktionsfähig, ggf Default Attributwerte VFS ergänzen

        self.buttons_value_k_nachbar_vfs["VFS 0"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS 1"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS 2"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS 3"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS 4"].SetValue(self.default_k_nachbar)
        self.buttons_value_k_nachbar_vfs["VFS 5"].SetValue(self.default_k_nachbar)

        self.buttons_value_n_versorger["VFS 0"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS 1"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS 2"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS 3"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS 4"].SetValue(self.default_anz_vf)
        self.buttons_value_n_versorger["VFS 5"].SetValue(self.default_anz_vf)

        self.cb_vfs.SetValue("TypeNo")
        self.cb_quelle.SetValue("None")
        self.cb_ziel.SetValue("None")
        self.cb_dist_fcn.SetValue("euclidean")

        self.SetStatusText('Default-Werte hergestellt')

    def event_choose_attr(self, event):
        attr = event.GetEventObject().GetStringSelection()

        if event.GetEventObject().Label == 'attr_vfs':
            self.attr_vfs = attr
            # Anpassen Buttons Attributwerte
            self.__set_values_vfs_buttons__()

        elif event.GetEventObject().Label == 'attr_quelle':
            if attr == 'None':
                self.attr_quelle = None
            else:
                self.attr_quelle = attr
        elif event.GetEventObject().Label == 'attr_ziel':
            if attr == 'None':
                self.attr_ziel = None
            else:
                self.attr_ziel = attr
        elif event.GetEventObject().Label == 'attr_dist_fcn':
            self.attr_dist_fcn = attr
        else:
            logging.warning("sollte nie passieren")

        # Erstellen einer neuen Calculator Instanz
        if self.visum is not None:
            # Init Calculator Instanz
            self.llt_calculator = llt.LuftlinienCalculator(self.visum,
                                                           attr_vfs=self.attr_vfs,
                                                           attr_quelle=self.attr_quelle,
                                                           attr_ziel=self.attr_ziel,
                                                           anz_versorger=1,
                                                           max_entfernung=1,
                                                           )
            # Übergebe aktuelle Parameter
            self.update_param_vfs()
        else:
            logging.warning("Umgang mit Nichtvisum Dateien ist nicht implementiert")

        self.SetStatusText("Attribut übernommen, Bezirke neu importiert, Rechnungen zurückgesetzt")

    def event_calculate(self, event):
        # Vorgehen
        # 1. Update der vorgegebenen parameter, falls was geändert wurde
        # 2. berechnen
        self.update_param_vfs()
        # self.llt_calculator.init_results() # bereits in calculate fcn implementiert
        self.llt_calculator.calculate_main()

        # Statusleiste
        self.SetStatusText('Berechnung durchgeführt')

    def event_quit_button(self, event):
        try:
            if self.visum is not None:
                self.visum = None  # Freigeben der Visum-Ressource
        except Exception as e:
            logging.error(f'Fehler beim Freigeben von Visum: {e}')

        # del self.llt_calculator
        self.tabLog.on_close()
        self.tabMain.Destroy()
        # del self.visum
        #self.stop = True
        self.panel.Destroy()
        # self.Close(True)
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
                                                           max_entfernung=1,
                                                           )
            # Übergebe aktuelle Parameter
            self.update_param_vfs()
        else:
            logging.warning("Umgang mit Nichtvisum Dateien ist nicht implementiert")
        self.SetStatusText('Daten importiert')

    def event_info(self, event):
        path_scripts = Path.cwd() #Path(self.visum.GetPath(37))
        logging.info(path_scripts)
        llt.show_info(path_scripts)

    def event_reset(self, event):
        if self.llt_calculator is None:
            a = 1  # todo
        else:
            self.llt_calculator.init_results()

        self.SetStatusText('Ergebnisse gelöscht, ggf Ergebnisse neu nach Visum importieren')

    def event_export_results(self, event):
        if self.llt_calculator is not None:
            self.llt_calculator.export_net(
                                           links_additive=True)
            self.llt_calculator.export_matrix()

    def event_export_net(self, event):
        vfs = event.GetEventObject().vfs

        if self.llt_calculator is not None:
            self.llt_calculator.export_net(list_vfs=[vfs],
                                           links_additive=True)
            self.llt_calculator.delete_unused_nodes()

        self.SetStatusText(f'{vfs}: Net-Datei exportiert und in Visum importiert')


    def event_export_mtx(self, event):
        vfs = event.GetEventObject().vfs

        if self.llt_calculator is not None:
            self.llt_calculator.export_matrix(list_vfs=[vfs])

        self.SetStatusText(f'{vfs}: Matrix in Visum geladen')

    def event_export_master(self, event):
        self.llt_calculator.export_matrix()
        self.llt_calculator.export_net(links_additive=True)
        self.llt_calculator.delete_unused_nodes()

        self.SetStatusText(f'die kombinierten Ergebnisse wurden in Visum importiert')

    def event_filter(self, event):
        if self.llt_calculator is not None:
            self.llt_calculator.filter_links_vfs()

    def event_delete_links(self, event):
        if self.llt_calculator is not None:
            self.llt_calculator.delete_added_links()
            self.llt_calculator.delete_unused_nodes()

    def update_param_vfs(self):
        if self.llt_calculator is not None:
            list_vfs = [vfs[0] for vfs in self.button_vfs_active.items() if vfs[1].Value > 0]
            dict_anz_versorger = {vfs: self.buttons_value_n_versorger[vfs].Value for vfs in list_vfs}
            dict_max_nachbar = {vfs: self.buttons_value_k_nachbar_vfs[vfs].Value for vfs in list_vfs}
            dict_vfs = {vfs: self.buttons_vfs_value[vfs].Value for vfs in list_vfs}

            self.llt_calculator.nachbarschaftsgrad_vfs = dict_max_nachbar
            self.llt_calculator.anz_versorger_vfs = dict_anz_versorger
            self.llt_calculator.vfs = dict_vfs
            self.llt_calculator.formula_dist = self.attr_dist_fcn

            logging.info(
f'''aktuelle Settings:
Bezirke: Attr. Zentralität-{self.llt_calculator.attr_central_level} Attr istQuelle-{self.llt_calculator.attr_is_from_zone} Attr istZiel-{self.llt_calculator.attr_is_to_zone} 
Distanzberechnung: {self.llt_calculator.formula_dist} 
VFS {self.llt_calculator.vfs} 
Nachbarschaftsgrad je VFS {self.llt_calculator.nachbarschaftsgrad_vfs} 
Anzahl Versorger je VFS {self.llt_calculator.anz_versorger_vfs}''')

## Spezifiziert & verwaltet den Tab mit den Eingabe- und Aktionsmöglichkeiten
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
                                  "VFS 1": wx.CheckBox(self, -1, "VFS 1"),
                                  "VFS 2": wx.CheckBox(self, -1, "VFS 2"),
                                  "VFS 3": wx.CheckBox(self, -1, "VFS 3"),
                                  "VFS 4": wx.CheckBox(self, -1, "VFS 4"),
                                  "VFS 5": wx.CheckBox(self, -1, "VFS 5")}

        tmp_iterator = 1
        for btn in self.button_vfs_active.values():
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 0), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.button_vfs_active = self.button_vfs_active

        # Spalte 2 Angabe Wert je VFS
        gridbagsizer1.Add(wx.StaticText(self, -1, "Attributwert VFS"),
                          pos=(0, 1), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_vfs_value = {"VFS 0": wx.SpinCtrl(self, -1, ""),
                                 "VFS 1": wx.SpinCtrl(self, -1, ""),
                                 "VFS 2": wx.SpinCtrl(self, -1, ""),
                                 "VFS 3": wx.SpinCtrl(self, -1, ""),
                                 "VFS 4": wx.SpinCtrl(self, -1, ""),
                                 "VFS 5": wx.SpinCtrl(self, -1, "")}
        tmp_iterator = 1
        for btn in self.buttons_vfs_value.values():
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 1), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_vfs_value = self.buttons_vfs_value

        # Spalte 2 Auswahl Austauschfunktion je VFS
        gridbagsizer1.Add(wx.StaticText(self, -1, "Austauschfunktion \n n-naechste Nachbarn"),
                          pos=(0, 2), flag=wx.ALIGN_CENTER | wx.ALL)

        self.buttons_value_k_nachbar_vfs = {"VFS 0": wx.SpinCtrl(self, -1, ""),
                                            "VFS 1": wx.SpinCtrl(self, -1, ""),
                                            "VFS 2": wx.SpinCtrl(self, -1, ""),
                                            "VFS 3": wx.SpinCtrl(self, -1, ""),
                                            "VFS 4": wx.SpinCtrl(self, -1, ""),
                                            "VFS 5": wx.SpinCtrl(self, -1, "")}
        tmp_iterator = 1
        for btn in self.buttons_value_k_nachbar_vfs.values():
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 2), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_value_k_nachbar_vfs = self.buttons_value_k_nachbar_vfs

        # Spalte 3 Versorgungsfunktion
        gridbagsizer1.Add(
            wx.StaticText(self, -1, "Versorgungsfunktion \n n Versorgungszentren"),
            pos=(0, 3), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_value_n_versorger = {"VFS 0": wx.SpinCtrl(self, -1, ""),
                                          "VFS 1": wx.SpinCtrl(self, -1, ""),
                                          "VFS 2": wx.SpinCtrl(self, -1, ""),
                                          "VFS 3": wx.SpinCtrl(self, -1, ""),
                                          "VFS 4": wx.SpinCtrl(self, -1, ""),
                                          "VFS 5": wx.SpinCtrl(self, -1, "")}

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
                                   "VFS 1": wx.Button(self, -1, "MTX"),
                                   "VFS 2": wx.Button(self, -1, "MTX"),
                                   "VFS 3": wx.Button(self, -1, "MTX"),
                                   "VFS 4": wx.Button(self, -1, "MTX"),
                                   "VFS 5": wx.Button(self, -1, "MTX")}

        tmp_iterator = 1
        for vfs, btn in self.buttons_export_mat.items():
            btn.vfs = vfs
            gridbagsizer1.Add(btn, pos=(tmp_iterator, 4), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        # Buttons Export Net
        self.buttons_export_net = {"VFS 0": wx.Button(self, -1, "Net"),
                                   "VFS 1": wx.Button(self, -1, "Net"),
                                   "VFS 2": wx.Button(self, -1, "Net"),
                                   "VFS 3": wx.Button(self, -1, "Net"),
                                   "VFS 4": wx.Button(self, -1, "Net"),
                                   "VFS 5": wx.Button(self, -1, "Net")}
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

        # Button Liste Distanzfkt
        self.cb_dist_fcn = wx.ComboBox(self, size=(200, -1),
                                             choices=["euclidean", "haversine"], style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)

        self.cb_dist_fcn.Label= 'attr_dist_fcn'
        self.TopLevelParent.cb_dist_fcn = self.cb_dist_fcn

        gridbagsizer1.Add(wx.StaticText(self, -1, "Funktion Distanzberechnung"), pos=(8, 0), span=(1, 1), flag=wx.EXPAND)
        gridbagsizer1.Add(self.cb_dist_fcn, pos=(8,1), span=(1,1), flag=wx.EXPAND)

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
        self.cb_dist_fcn.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)

## Spezifiziert den Tab, der die Lognachrichten ausgibt
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

        # # Bind the panel destruction event to ensure cleanup
        # self.Bind(wx.EVT_WINDOW_DESTROY, self.on_close)

    def on_close(self):
        """ Perform any cleanup actions here """

        # Remove and close the custom handler
        self.logger.removeHandler(self.handler)
        self.handler.close()

        # Remove the file and stream handlers (optional if they need cleanup)
        for handler in self.logger.handlers[:]:
            if isinstance(handler, (logging.FileHandler, logging.StreamHandler)):
                self.logger.removeHandler(handler)
                handler.close()

        # event.Skip()  # Ensure the event propagates to the parent if needed

    def __del__(self):
        """ Destructor, ensure the logger handler is removed """
        self.logger.removeHandler(self.handler)

## Handler der Logbefehle
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
