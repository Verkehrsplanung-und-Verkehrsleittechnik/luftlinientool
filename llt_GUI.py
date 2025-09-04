import wx
import wx.html2
import luftlinientool as llt
from pathlib import Path
import logging
from language_management import Translator

# ===== Helper Functions =====
## Loads all Visum zone attributes.
#  @param Visum The Visum instance to get attributes from.
#  @return A list of all zone attribute IDs.
def get_attr_zones(Visum):
    list_attr = Visum.Net.Zones.Attributes.GetAll
    list_attr_id = [attr.ID for attr in list_attr]

    return list_attr_id


# ======= Classes ======

## @class LLTFrame
#  @brief Defines the complete window, creates the individual components and connects them with the logic.
#
#  The `LLTFrame` class is the main window of the application. It creates and manages all UI components,
#  handles user interactions, and connects the UI with the underlying logic.
class LLTFrame(wx.Frame):
    ## Initializes the main application window.
    #  @param translator The translator object to handle language localization.
    def __init__(self, translator):
        super().__init__(parent=None)

        self.translator = translator # Pointer to translator

        # ===== Attribute =====
        self.buttons_value_n_supplier = None
        self.cb_origin = None # Renamed from cb_quelle
        self.cb_destination = None # Renamed from cb_ziel
        self.cb_cfl = None
        self.buttons_cfl_value = None
        self.buttons_value_k_neighbor_cfl = None
        self.button_cfl_active = None
        self.llt_calculator = None  # llt.LuftlinienCalculator()
        self.default_k_neighbor = 1
        self.default_no_supplier = 0
        self.attr_cfl = "TypeNo"
        self.attr_origin = None
        self.attr_destination = None # Renamed from attr_ziel

        # If Visum exists -> do nothing
        # otherwise open a window to select a file
        try:
            # tests if the variable Visum exists
            global Visum
            Visum
        except NameError:
            import win32com.client as com
            defDir = Path.cwd()
            with wx.FileDialog(self,
                               self.translator.translate('file_selection'),
                               str(defDir),
                               wildcard=(self.translator.translate('ver_files')+ " (*.ver)|*.ver"), # Use self.translator
                               style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as file_dlg:
                if file_dlg.ShowModal() == wx.ID_CANCEL:
                    return  # the user changed their mind
                source = Path(file_dlg.Path)

            if source.suffix == ".ver":
                version = 250
                Visum = com.Dispatch(f"Visum.Visum.{version}")
                logging.info('open visum file: {}'.format(source))
                Visum.LoadVersion(source)

        self.visum = Visum
        self.list_attr = get_attr_zones(self.visum)

        self.attr_origin = None
        self.attr_destination = None # Renamed from attr_ziel
        self.attr_cfl = "TypeNo"
        self.attr_dist_fcn = "euclidean"

        self.__set_layout__()
        self.__set_properties__()
        self.__bind_events__()

        self.Show()

    ## Sets the properties of the main window.
    #  Configures window title, size, and initializes default values.
    def __set_properties__(self):
        self.SetTitle(self.translator.translate('app_title')) # Use self.translator
        self.SetMinSize((1200, 500))
        # Set the initial size of the window to be larger
        self.SetSize((1400, 500)) # Added this line to set the initial window size

        self.__set_values_cfl_buttons__()

        self.event_set_default()

    ## Sets up the layout of the main window.
    #  Creates and arranges all UI components including panels, notebook, menu bar, and status bar.
    def __set_layout__(self):
        # Layout
        # Menu bar/Toolbar at the top
        # then

        # create a panel (between menubar and statusbar) ...
        self.panel = wx.Panel(self)
        self.notebook = wx.Notebook(self.panel)

        # create tabs
        self.tabMain = MainTab(self.notebook, self.translator)
        self.tabLog = LogTab(self.notebook)

        self.notebook.AddPage(self.tabMain, "Main")
        self.notebook.AddPage(self.tabLog, "Log")

        # create a menubar at the top of the user frame
        self.menu_bar = wx.MenuBar()

        # create a menu ...
        self.menu = wx.Menu()
        self.menu.Append(10, self.translator.translate("language_choice"))
        self.menu.AppendSeparator()
        self.menu.Append(11, self.translator.translate('menu_import_data'))
        self.menu.Append(12, self.translator.translate('menu_calculate'))
        self.menu.AppendSeparator()
        self.menu.Append(13, self.translator.translate('menu_reset_calculations'))
        self.menu.Append(14, self.translator.translate('menu_set_defaults'))
        self.menu.AppendSeparator()
        self.menu.Append(15, "&Info")

        self.menu.AppendSeparator()
        # put the menu on the menubar
        self.menu_bar.Append(self.menu, self.translator.translate('options_tab'))
        self.SetMenuBar(self.menu_bar)
        self.toolbar = self.CreateToolBar(style=wx.TB_TEXT | wx.TB_NOICONS)

        # Workaround keine Bilder zur Verfügung: Leeres Bitmap Objekt
        self.toolbar.AddTool(100, self.translator.translate('language_choice'), wx.Bitmap())
        self.toolbar.AddTool(101, self.translator.translate('menu_import_data'), wx.Bitmap())
        self.toolbar.AddTool(102, self.translator.translate('menu_calculate'), wx.Bitmap())
        self.toolbar.AddTool(103, self.translator.translate('menu_reset_calculations'), wx.Bitmap())
        self.toolbar.AddTool(104, self.translator.translate('default_Values_tab'), wx.Bitmap())
        self.toolbar.AddTool(105, 'Info', wx.Bitmap())
        self.toolbar.AddTool(106, self.translator.translate('toolbar_filter_inserted_links'), wx.Bitmap())
        self.toolbar.AddTool(107, self.translator.translate('toolbar_delete_inserted_links'), wx.Bitmap())
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

    ## Binds event handlers to UI elements.
    # Sets up all event bindings for menu items, toolbar buttons, and other UI elements.
    def __bind_events__(self):
        # Event Handler
        # bind the menu event to an event handler, share QuitBtn event
        self.Bind(wx.EVT_CLOSE, self.event_quit_button)
        self.Bind(wx.EVT_MENU, self.on_choose_language, id=10)
        self.Bind(wx.EVT_MENU, self.event_import_data, id=11)
        self.Bind(wx.EVT_MENU, self.event_calculate, id=12)
        self.Bind(wx.EVT_MENU, self.event_info, id=15)
        self.Bind(wx.EVT_MENU, self.event_reset, id=13)
        self.Bind(wx.EVT_MENU, self.event_set_default, id=14)

        self.toolbar.Bind(wx.EVT_TOOL, self.on_choose_language, id=100)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_import_data, id=101)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_calculate, id=102)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_reset, id=103)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_set_default, id=104)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_info, id=105)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_filter, id=106)
        self.toolbar.Bind(wx.EVT_TOOL, self.event_delete_links, id=107)

    ## Sets the values and ranges for the CFL buttons.
    #  Configures the range and initial values for the CFL value buttons based on the maximum value of the selected attribute.
    def __set_values_cfl_buttons__(self):
        max_value_cfl = int(self.visum.Net.AttValue(f"Max:Zones\{self.attr_cfl}"))
        idx = 0
        for btn in self.buttons_cfl_value.values():
            btn.SetRange(0, max_value_cfl)
            btn.SetValue(idx)
            idx += 1

    ## Event handler for language selection.
    #  Opens a dialog for the user to select a language, updates the translator with the selected language,
    #  and refreshes the GUI text to reflect the new language.
    #  @param event The event object.
    def on_choose_language(self, event):

        languages = sorted(list(self.translator.translations.keys()))

        # current_selection_index = languages.index(self.translator.get_selected_language())

        # Erstelle den wx.SingleChoiceDialog
        dlg = wx.SingleChoiceDialog(
            self,
            self.translator.translate('dialog_language_select_message'),
            self.translator.translate('dialog_language_select_title'),
            choices=languages,
        )

        if dlg.ShowModal() == wx.ID_OK:
            selected_language = languages[dlg.GetSelection()]

            if selected_language != self.translator.get_selected_language():
                # Change the language in the translator
                self.translator.update_selected_language(selected_language)
                # Update the GUI texts (Step 3)
                self.refresh_gui_text()

        dlg.Destroy()  # Important: Destroy the dialog after it has been closed.

    ## Event handler for setting default values.
    #  Resets all input fields to their default values, including neighbor counts,
    #  supplier counts, and attribute selections.
    #  @param event The event object (optional).
    def event_set_default(self, event=None):
        # functional, may need to add default attribute values for CFL

        self.buttons_value_k_neighbor_cfl['VFS 0'].SetValue(self.default_k_neighbor) # Reverted to 'VFS 0'
        self.buttons_value_k_neighbor_cfl['VFS 1'].SetValue(self.default_k_neighbor)
        self.buttons_value_k_neighbor_cfl['VFS 2'].SetValue(self.default_k_neighbor)
        self.buttons_value_k_neighbor_cfl['VFS 3'].SetValue(self.default_k_neighbor)
        self.buttons_value_k_neighbor_cfl['VFS 4'].SetValue(self.default_k_neighbor)
        self.buttons_value_k_neighbor_cfl['VFS 5'].SetValue(self.default_k_neighbor)

        self.buttons_value_n_supplier['VFS 0'].SetValue(self.default_no_supplier) # Reverted to 'VFS 0'
        self.buttons_value_n_supplier['VFS 1'].SetValue(self.default_no_supplier)
        self.buttons_value_n_supplier['VFS 2'].SetValue(self.default_no_supplier)
        self.buttons_value_n_supplier['VFS 3'].SetValue(self.default_no_supplier)
        self.buttons_value_n_supplier['VFS 4'].SetValue(self.default_no_supplier)
        self.buttons_value_n_supplier['VFS 5'].SetValue(self.default_no_supplier)

        self.cb_cfl.SetValue("TypeNo")
        self.cb_origin.SetValue("None")
        self.cb_destination.SetValue("None")
        self.cb_dist_fcn.SetValue("euclidean")

        self.SetStatusText(self.translator.translate('set_defaults'))

    ## Event handler for attribute selection changes.
    #  Updates the appropriate attribute based on which combo box triggered the event,
    #  reinitializes the calculator with the new attributes, and updates the status bar.
    #  @param event The event object.
    def event_choose_attr(self, event):
        attr = event.GetEventObject().GetStringSelection()

        if event.GetEventObject().Label == self.translator.translate('attr_cfl'):
            self.attr_cfl = attr
            # Adjust buttons for attribute values
            self.__set_values_cfl_buttons__()

        elif event.GetEventObject().Label == self.translator.translate('attr_origin'):
            if attr == 'None':
                self.attr_origin = None
            else:
                self.attr_origin = attr
        elif event.GetEventObject().Label == self.translator.translate('attr_Destination'):
            if attr == 'None':
                self.attr_destination = None
            else:
                self.attr_destination = attr
        elif event.GetEventObject().Label == 'attr_dist_fcn':
            self.attr_dist_fcn = attr
        else:
            logging.warning(self.translator.translate('warning_should_never_happen'))

        # Creating a new Calculator instance
        if self.visum is not None:
            # Initialize Calculator instance
            self.llt_calculator = llt.LuftlinienCalculator(self.visum, attr_cfl=self.attr_cfl, max_distance=1,
                                                           no_suppliers=1, attr_orig=self.attr_origin,
                                                           attr_dest=self.attr_destination,translator=self.translator)
            # Pass current parameters
            self.update_param_cfl()
        else:
            logging.warning(self.translator.translate('warning_non_visum_files_not_supported'))
        self.SetStatusText(self.translator.translate('status_attribute_applied_reimport_reset'))

    ## Event handler for the calculate button.
    #  Updates the parameters from the UI, performs the main calculation,
    #  and updates the status bar to indicate completion.
    #  @param event The event object.
    def event_calculate(self, event):
        # Procedure
        # 1. Update the specified parameters if something has been changed
        # 2. Calculate
        self.update_param_cfl()
        # self.llt_calculator.init_results() # already implemented in calculate function
        self.llt_calculator.calculate_main()

        # Status bar
        self.SetStatusText(self.translator.translate('calculation_Performed'))

    ## Event handler for the quit button or window close event.
    #  Destroys the panel and frame, and exits the application's main loop.
    #  @param event The event object.
    def event_quit_button(self, event):
        # del self.visum
        # self.stop = True
        self.panel.Destroy()
        self.Destroy()
        wx.GetApp().ExitMainLoop()

    ## Event handler for importing data.
    #  Creates a new LuftlinienCalculator instance with the current attributes,
    #  updates parameters, and updates the status bar.
    #  @param event The event object.
    def event_import_data(self, event):
        # works so far,

        # Creating a Calculator instance
        if self.visum is not None:
            # Init Calculator instance
            self.llt_calculator = llt.LuftlinienCalculator(self.visum, attr_cfl=self.attr_cfl, max_distance=1,
                                                           no_suppliers=1, attr_orig=self.attr_origin,
                                                           attr_dest=self.attr_destination,translator=self.translator)
            # Pass current parameters
            self.update_param_cfl()
        else:
            logging.warning(self.translator.translate('warning_non_visum_files_not_supported'))
        self.SetStatusText(self.translator.translate('data_Imported'))

    ## Event handler for the Info button.
    #  @param event The event object.
    def event_info(self, event):
        path_scripts = Path.cwd() # Path(self.visum.GetPath(37))
        logging.info(path_scripts)
        # Update this line to instantiate HelpPopUp directly and pass the translator.
        HelpPopUp(self, self.translator, path_scripts)

    ## Event handler for resetting calculations.
    #  Initializes the results in the calculator, effectively clearing any previous calculations,
    #  and updates the status bar.
    #  @param event The event object.
    def event_reset(self, event):
        if self.llt_calculator is None:
            a = 1  # todo
        else:
            self.llt_calculator.init_results()

        self.SetStatusText(self.translator.translate('status_results_deleted_may_reimport'))

    ## Event handler for exporting all results.
    #  Exports both the network and matrix results from the calculator if it exists.
    #  @param event The event object.
    def event_export_results(self, event):
        if self.llt_calculator is not None:
            self.llt_calculator.export_net(links_additive=True)
            self.llt_calculator.export_matrix()

    ## Event handler for exporting network results for a specific CFL level.
    #  Exports the network results for the specified CFL level, deletes unused nodes,
    #  and updates the status bar with the CFL level that was exported.
    #  @param event The event object containing the CFL level to export.
    def event_export_net(self, event):
        cfl_level = event.GetEventObject().cfl

        if self.llt_calculator is not None:
            self.llt_calculator.export_net(links_additive=True, list_cfl=[cfl_level])
            self.llt_calculator.delete_unused_nodes()

        self.SetStatusText(f'{cfl_level}'+self.translator.translate(': status_net_exported_imported'))

    ## Event handler for exporting matrix results for a specific CFL level.
    #  Exports the matrix results for the specified CFL level and updates the status bar
    #  with the CFL level that was exported.
    #  @param event The event object containing the CFL level to export.
    def event_export_mtx(self, event):
        cfl_level = event.GetEventObject().cfl

        if self.llt_calculator is not None:
            self.llt_calculator.export_matrix(list_cfl=[cfl_level])  # list_vfs parameter name remains as per llt.py

        self.SetStatusText(f'{cfl_level}'+self.translator.translate(': status_matrix_loaded_into_visum'))

    ## Event handler for exporting all results at once.
    #  Exports both matrix and network results for all CFL levels, deletes unused nodes,
    #  and updates the status bar.
    #  @param event The event object.
    def event_export_master(self, event):
        self.llt_calculator.export_matrix()
        self.llt_calculator.export_net(links_additive=True)
        self.llt_calculator.delete_unused_nodes()

        self.SetStatusText(self.translator.translate(f'status_combined_results_imported'))

    ## Event handler for filtering links.
    #  Applies a filter to show only the links created by the calculator.
    #  @param event The event object.
    def event_filter(self, event):
        if self.llt_calculator is not None:
            self.llt_calculator.filter_links_cfl()

    ## Event handler for deleting links.
    #  Deletes all links created by the calculator and removes any unused nodes.
    #  @param event The event object.
    def event_delete_links(self, event):
        if self.llt_calculator is not None:
            self.llt_calculator.delete_added_links()
            self.llt_calculator.delete_unused_nodes()

    ## Updates the calculator parameters from the UI.
    #  Collects the current values from the UI controls and updates the calculator's parameters,
    #  including CFL levels, supplier counts, neighbor counts, and distance formula. Also logs the current settings.
    def update_param_cfl(self):
        if self.llt_calculator is not None:
            list_cfl_levels = [cfl_level[0] for cfl_level in self.button_cfl_active.items() if cfl_level[1].Value > 0]
            dict_num_suppliers = {cfl_level: self.buttons_value_n_supplier[cfl_level].Value for cfl_level in list_cfl_levels}
            dict_max_neighbor = {cfl_level: self.buttons_value_k_neighbor_cfl[cfl_level].Value for cfl_level in list_cfl_levels}
            dict_cfl_values = {cfl_level: self.buttons_cfl_value[cfl_level].Value for cfl_level in list_cfl_levels}

            self.llt_calculator.max_neighbor_cfl = dict_max_neighbor
            self.llt_calculator.num_suppliers_cfl = dict_num_suppliers
            self.llt_calculator.cfl = dict_cfl_values
            self.llt_calculator.formula_dist = self.attr_dist_fcn

            logging.info( self.translator.translate('Current settings:') + "\n" +
                          self.translator.translate('Districts: Attr. Centrality')+f'{self.llt_calculator.attr_central_level}' + "\n"+
                          self.translator.translate('label_district_is_origin') + f'{self.llt_calculator.attr_is_from_zone}' + "\n" +
                          self.translator.translate('label_district_is_destination') + f'{self.llt_calculator.attr_is_to_zone}' + "\n" +
                          self.translator.translate('distance_calculation_Setting') + f'{self.llt_calculator.formula_dist}' + "\n" +
                          self.translator.translate('CFL') + f'{self.llt_calculator.cfl}' + "\n" +
                          self.translator.translate('neighbourhood_LevelPerCFL_Setting') + f'{self.llt_calculator.max_neighbor_cfl}' + "\n" +
                          self.translator.translate('number_of_Suppliers_PerCFL_Setting') + f'{self.llt_calculator.num_suppliers_cfl}')


    ## Updates all text elements in the GUI to the current language.
    #  Also triggers refresh_gui_text on child components.
    def refresh_gui_text(self):
        # 1. Hauptfenstertitel aktualisieren
        self.SetTitle(self.translator.translate('app_title'))

        # 2. Notebook-Tab-Titel aktualisieren
        self.notebook.SetPageText(0, self.translator.translate('Main tab'))
        self.notebook.SetPageText(1, self.translator.translate('Log tab'))

        # 3. Menüleiste aktualisieren
        self.menu.FindItemById(10).SetItemLabel(self.translator.translate("language_choice"))
        self.menu.FindItemById(11).SetItemLabel(self.translator.translate('menu_import_data'))
        self.menu.FindItemById(12).SetItemLabel(self.translator.translate('menu_calculate'))
        self.menu.FindItemById(13).SetItemLabel(self.translator.translate('menu_reset_calculations'))
        self.menu.FindItemById(14).SetItemLabel(self.translator.translate('menu_set_defaults'))
        self.menu.FindItemById(15).SetItemLabel(self.translator.translate('Info'))
        # Update the menu bar's overall menu label using its index (assuming it's the first menu added, index 0)
        self.menu_bar.SetMenuLabel(0, self.translator.translate('options_tab'))


        # 4. Toolbar aktualisieren
        self.toolbar.FindById(100).SetLabel(self.translator.translate('language_choice'))
        self.toolbar.FindById(101).SetLabel(self.translator.translate('menu_import_data'))
        self.toolbar.FindById(102).SetLabel(self.translator.translate('menu_calculate'))
        self.toolbar.FindById(103).SetLabel(self.translator.translate('menu_reset_calculations'))
        self.toolbar.FindById(104).SetLabel(self.translator.translate('default_Values_tab'))
        self.toolbar.FindById(105).SetLabel(self.translator.translate('Info'))
        self.toolbar.FindById(106).SetLabel(self.translator.translate('toolbar_filter_inserted_links'))
        self.toolbar.FindById(107).SetLabel(self.translator.translate('toolbar_delete_inserted_links'))

        # 5. Unterkomponenten (Tabs) aktualisieren
        self.tabMain.refresh_gui_text()
        self.tabLog.refresh_gui_text()

        # Update the status text
        self.SetStatusText(self.translator.translate('set_defaults'))

        # Wichtig: Layout und Refresh erzwingen nach Textänderungen
        self.Layout()
        self.Refresh()


## @class MainTab
#  @brief Specifies and manages the tab with input and action options.
#
#  The `MainTab` class creates and manages the main tab of the application, which contains
#  all the input fields, buttons, and other controls for configuring and executing the
#  air-line calculations.
class MainTab(wx.Panel):
    ## Initializes the main tab panel.
    #  @param parent The parent window that contains this panel.
    #  @param translator The translator object to handle language localization.
    def __init__(self, parent, translator):
        wx.Panel.__init__(self, parent)

        self.translator = translator
        # Initialize instance attributes for widgets and sizers that need to be updated
        self.hbox1 = None
        self.gridbagsizer1 = None
        self.static_text_centrality = None
        self.static_text_origin = None
        self.static_text_destination = None
        self.static_text_cfl_label = None
        self.static_text_attr_cfl = None
        self.static_text_exchange_fcn = None
        self.static_text_supply_fcn = None
        self.static_text_visum_as = None
        self.static_text_dist_fcn_label = None
        self.btn_export_master = None # Ensure it's initialized

        self.__set_layout__()
        self.__bind_events__()

    ## Sets up the layout of the main tab panel.
    #  Creates and arranges all UI components including district attribute selection,
    #  connectivity function level parameters, and action buttons.
    def __set_layout__(self):
        # Rows with individual elements (vbox_outer)
        # Row 1: District attribute selection
        # Row 2: GridbagSizer with everything except Log
        # Status bar at the bottom

        vbox_outer = wx.BoxSizer(wx.VERTICAL)
        self.hbox1 = wx.BoxSizer(wx.HORIZONTAL) # Make hbox1 an instance attribute
        self.gridbagsizer1 = wx.GridBagSizer(vgap=10, hgap=50) # Make gridbagsizer1 an instance attribute

        # Selection of district attributes
        self.cb_cfl= wx.ComboBox(self, size=(200, -1), choices=self.TopLevelParent.list_attr,
                                 style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)
        self.cb_cfl.Label = self.translator.translate('attr_cfl')
        self.TopLevelParent.cb_cfl = self.cb_cfl

        self.cb_origin = wx.ComboBox(self, size=(200, -1), choices=self.TopLevelParent.list_attr,
                                     style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)
        self.cb_origin.Label = self.translator.translate('attr_origin')
        self.TopLevelParent.cb_origin = self.cb_origin

        self.cb_destination = wx.ComboBox(self, size=(200, -1), choices=self.TopLevelParent.list_attr,
                                   style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)
        self.cb_destination.Label = self.translator.translate('attr_destination')
        self.TopLevelParent.cb_destination = self.cb_destination

        # Store StaticText widgets as instance attributes
        self.static_text_centrality = wx.StaticText(self, -1, (self.translator.translate('district_Attribute_Setting')+ "\n"+self.translator.translate('centrality_Setting')))
        self.hbox1.Add(self.static_text_centrality, 0, wx.ALL | wx.EXPAND, 5)
        self.hbox1.Add(self.cb_cfl, 0, wx.ALL | wx.EXPAND, 15)

        self.static_text_origin = wx.StaticText(self, -1, (self.translator.translate('district_Attribute_Setting')+ "\n"+self.translator.translate('isOrigin_Option')))
        self.hbox1.Add(self.static_text_origin, 0, wx.ALL | wx.EXPAND, 5)
        self.hbox1.Add(self.cb_origin, 0, wx.ALL | wx.EXPAND, 15)

        self.static_text_destination = wx.StaticText(self, -1, (self.translator.translate('district_Attribute_Setting')+ "\n"+self.translator.translate('isDestination_Option')))
        self.hbox1.Add(self.static_text_destination, 0, wx.ALL | wx.EXPAND, 5)
        self.hbox1.Add(self.cb_destination, 0, wx.ALL | wx.EXPAND, 15)

        # Header Column 1
        self.static_text_cfl_label = wx.StaticText(self, -1, self.translator.translate('connectivity_Function_Level_Parameter'))
        self.gridbagsizer1.Add(self.static_text_cfl_label,
                          pos=(0, 0), flag=wx.TOP | wx.LEFT | wx.BOTTOM, border=5)

        self.button_cfl_active = {"VFS 0": wx.CheckBox(self, -1, self.translator.translate('VFS 0')), # Reverted to 'VFS 0'
                                  "VFS 1": wx.CheckBox(self, -1, self.translator.translate('VFS 1')),
                                  "VFS 2": wx.CheckBox(self, -1, self.translator.translate('VFS 2')),
                                  "VFS 3": wx.CheckBox(self, -1, self.translator.translate('VFS 3')),
                                  "VFS 4": wx.CheckBox(self, -1, self.translator.translate('VFS 4')),
                                  "VFS 5": wx.CheckBox(self, -1, self.translator.translate('VFS 5'))}

        tmp_iterator = 1
        for btn in self.button_cfl_active.values():
            self.gridbagsizer1.Add(btn, pos=(tmp_iterator, 0), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.button_cfl_active = self.button_cfl_active

        # Column 2 Value specification per CFL
        self.static_text_attr_cfl = wx.StaticText(self, -1, self.translator.translate('attribute_value_CFL_Parameter')) # Reverted to 'Attributwert VFS'
        self.gridbagsizer1.Add(self.static_text_attr_cfl,
                          pos=(0, 1), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_cfl_value = {"VFS 0": wx.SpinCtrl(self, -1, ""), # Reverted to 'VFS 0'
                                  "VFS 1": wx.SpinCtrl(self, -1, ""),
                                  "VFS 2": wx.SpinCtrl(self, -1, ""),
                                  "VFS 3": wx.SpinCtrl(self, -1, ""),
                                  "VFS 4": wx.SpinCtrl(self, -1, ""),
                                  "VFS 5": wx.SpinCtrl(self, -1, "")}
        tmp_iterator = 1
        for btn in self.buttons_cfl_value.values():
            self.gridbagsizer1.Add(btn, pos=(tmp_iterator, 1), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_cfl_value = self.buttons_cfl_value

        # Column 2 Selection of exchange function per CFL
        self.static_text_exchange_fcn = wx.StaticText(self, -1, (self.translator.translate('interchangeFunction_Parameter')+ "\n"  + self.translator.translate('n_Nearest_Neighbour_Parameter')))
        self.gridbagsizer1.Add(self.static_text_exchange_fcn,
                          pos=(0, 2), flag=wx.ALIGN_CENTER | wx.ALL)

        self.buttons_value_k_neighbor_cfl = {"VFS 0": wx.SpinCtrl(self, -1, ""), # Reverted to 'VFS 0'
                                            "VFS 1": wx.SpinCtrl(self, -1, ""),
                                            "VFS 2": wx.SpinCtrl(self, -1, ""),
                                            "VFS 3": wx.SpinCtrl(self, -1, ""),
                                            "VFS 4": wx.SpinCtrl(self, -1, ""),
                                            "VFS 5": wx.SpinCtrl(self, -1, "")}
        tmp_iterator = 1
        for btn in self.buttons_value_k_neighbor_cfl.values():
            self.gridbagsizer1.Add(btn, pos=(tmp_iterator, 2), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_value_k_neighbor_cfl = self.buttons_value_k_neighbor_cfl

        # Column 3 Supply function
        self.static_text_supply_fcn = wx.StaticText(self, -1, (self.translator.translate('supply_Function_Parameter')+"\n"+self.translator.translate('n_Supply_Centers_Parameter')))
        self.gridbagsizer1.Add(self.static_text_supply_fcn,
            pos=(0, 3), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_value_n_supplier = {"VFS 0": wx.SpinCtrl(self, -1, ""), # Reverted to 'VFS 0'
                                          "VFS 1": wx.SpinCtrl(self, -1, ""),
                                          "VFS 2": wx.SpinCtrl(self, -1, ""),
                                          "VFS 3": wx.SpinCtrl(self, -1, ""),
                                          "VFS 4": wx.SpinCtrl(self, -1, ""),
                                          "VFS 5": wx.SpinCtrl(self, -1, "")}

        tmp_iterator = 1
        for btn in self.buttons_value_n_supplier.values():
            self.gridbagsizer1.Add(btn, pos=(tmp_iterator, 3), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        self.TopLevelParent.buttons_value_n_supplier = self.buttons_value_n_supplier

        # Export Matrix Buttons
        self.static_text_visum_as = wx.StaticText(self, -1, self.translator.translate('create_In_VisumAs_Option'))
        self.gridbagsizer1.Add(self.static_text_visum_as,
            pos=(0, 4), span=(1, 2), flag=wx.ALIGN_CENTER | wx.ALL)
        self.buttons_export_mat = {"VFS 0": wx.Button(self, -1, "MTX"), # Reverted to 'VFS 0'
                                   "VFS 1": wx.Button(self, -1, "MTX"),
                                   "VFS 2": wx.Button(self, -1, "MTX"),
                                   "VFS 3": wx.Button(self, -1, "MTX"),
                                   "VFS 4": wx.Button(self, -1, "MTX"),
                                   "VFS 5": wx.Button(self, -1, "MTX")}

        tmp_iterator = 1
        for cfl_level, btn in self.buttons_export_mat.items():
            btn.cfl = cfl_level
            self.gridbagsizer1.Add(btn, pos=(tmp_iterator, 4), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        # Buttons Export Net
        self.buttons_export_net = {"VFS 0": wx.Button(self, -1, "Net"), # Reverted to 'VFS 0'
                                   "VFS 1": wx.Button(self, -1, "Net"),
                                   "VFS 2": wx.Button(self, -1, "Net"),
                                   "VFS 3": wx.Button(self, -1, "Net"),
                                   "VFS 4": wx.Button(self, -1, "Net"),
                                   "VFS 5": wx.Button(self, -1, "Net")}
        tmp_iterator = 1
        for cfl_level, btn in self.buttons_export_net.items():
            btn.cfl = cfl_level
            self.gridbagsizer1.Add(btn, pos=(tmp_iterator, 5), flag=wx.ALIGN_CENTER)
            tmp_iterator += 1

        # Buttons export all
        self.btn_export_master = wx.Button(self, -1, (self.translator.translate('import_To_Visum_All_CFL_Option') +"\n"+self.translator.translate('routes_And_Matrix_Option'))) # Reverted to 'Import nach Visum alle VFS'
        self.btn_export_master.cfl_level = 'alle'
        self.gridbagsizer1.Add(self.btn_export_master,
                          pos=(7, 4), span=(3, 2), flag=wx.EXPAND)

        # Button Liste Distanzfkt
        self.cb_dist_fcn = wx.ComboBox(self, size=(200, -1),
                                       choices=["euclidean", "haversine"],
                                       style=wx.CB_DROPDOWN | wx.CB_READONLY | wx.CB_SORT)

        self.cb_dist_fcn.Label = 'attr_dist_fcn'
        self.TopLevelParent.cb_dist_fcn = self.cb_dist_fcn

        self.static_text_dist_fcn_label = wx.StaticText(self, -1, self.translator.translate('distance_Calculation_Fcn_Option'))
        self.gridbagsizer1.Add(self.static_text_dist_fcn_label, pos=(8, 0), span=(1, 1),
                          flag=wx.EXPAND)
        self.gridbagsizer1.Add(self.cb_dist_fcn, pos=(8, 1), span=(1, 1), flag=wx.EXPAND)

        # Aufbau Layout
        vbox_outer.Add(self.hbox1, 0, wx.ALL | wx.EXPAND, 1)
        vbox_outer.Add(self.gridbagsizer1, 1, wx.ALL | wx.EXPAND, 6)
        self.SetSizer(vbox_outer)

        # ==== Event binding

    ## Binds event handlers to UI elements in the main tab.
    #  Sets up all event bindings for buttons and combo boxes in the main tab,
    #  connecting them to the appropriate event handlers in the parent frame.
    def __bind_events__(self):

        for cfl_level, btn in self.buttons_export_net.items():
            btn.Bind(wx.EVT_BUTTON, self.TopLevelParent.event_export_net)

        for cfl_level, btn in self.buttons_export_mat.items():
            btn.Bind(wx.EVT_BUTTON, self.TopLevelParent.event_export_mtx)

        self.btn_export_master.Bind(wx.EVT_BUTTON, self.TopLevelParent.event_export_master)

        self.cb_cfl.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)
        self.cb_origin.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)
        self.cb_destination.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)
        self.cb_dist_fcn.Bind(wx.EVT_COMBOBOX, self.TopLevelParent.event_choose_attr)


    ## Updates all text elements in the main tab to the current language.
    def refresh_gui_text(self):
        # Update ComboBox labels
        self.cb_cfl.Label = self.translator.translate('attr_cfl')
        self.cb_origin.Label = self.translator.translate('attr_origin')
        self.cb_destination.Label = self.translator.translate('attr_destination')

        # Update StaticText widgets
        self.static_text_centrality.SetLabel(self.translator.translate('district_Attribute_Setting')+ "\n"+self.translator.translate('centrality_Setting'))
        self.static_text_origin.SetLabel(self.translator.translate('district_Attribute_Setting')+ "\n"+self.translator.translate('isOrigin_Option'))
        self.static_text_destination.SetLabel(self.translator.translate('district_Attribute_Setting')+ "\n"+self.translator.translate('isDestination_Option'))
        self.static_text_cfl_label.SetLabel(self.translator.translate('connectivity_Function_Level_Parameter'))
        self.static_text_attr_cfl.SetLabel(self.translator.translate('attribute_value_CFL_Parameter')) # Reverted to 'Attributwert VFS'
        self.static_text_exchange_fcn.SetLabel(self.translator.translate('interchangeFunction_Parameter') + "\n" + self.translator.translate('n_Nearest_Neighbour_Parameter'))
        self.static_text_supply_fcn.SetLabel(self.translator.translate('supply_Function_Parameter') + "\n" + self.translator.translate('n_Supply_Centers_Parameter'))
        self.static_text_visum_as.SetLabel(self.translator.translate('create_In_VisumAs_Option'))
        self.static_text_dist_fcn_label.SetLabel(self.translator.translate('distance_Calculation_Fcn_Option'))

        # Update CheckBox labels (iterate over existing objects)
        for key, checkbox in self.button_cfl_active.items():
            # The keys are 'VFS 0', 'VFS 1', etc., so we translate these keys directly
            checkbox.SetLabel(self.translator.translate(key))

        # Update master export button label
        self.btn_export_master.SetLabel(self.translator.translate('import_To_Visum_All_CFL_Option') +"\n"+self.translator.translate('routes_And_Matrix_Option')) # Reverted to 'Import nach Visum alle VFS'

        self.Layout()
        self.Refresh()


## @class LogTab
#  @brief Specifies the tab that displays log messages.
#
#  The `LogTab` class creates and manages the log tab of the application, which displays
#  log messages from the application. It sets up logging handlers to capture and display
#  messages in a text control.
class LogTab(wx.Panel):
    ## Initializes the log tab panel.
    #  @param parent The parent window that contains this panel.
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(wx.StaticText(self, -1, "Message-Log"), 0, wx.ALL | wx.CENTER, 5)
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

        self.log = wx.TextCtrl(self, wx.ID_ANY, size=(700, 200),
                               style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL | wx.EXPAND)
        self.handler = WxTextCtrlHandler(self.log)
        self.handler.setFormatter(logger_format)
        self.logger.addHandler(self.handler)

        vbox.Add(self.log, 1, wx.ALL | wx.EXPAND, 5)
        self.SetSizer(vbox)

        # Bind the panel destruction event to ensure cleanup
        self.Bind(wx.EVT_WINDOW_DESTROY, self.on_close)

    ## Handles the window destruction event.
    #  Performs cleanup actions when the panel is destroyed, removing and closing log handlers.
    #  @param event The window destruction event.
    def on_close(self, event):
        """ Perform any cleanup actions here """

        # Remove and close the custom handler
        self.logger.removeHandler(self.handler)
        self.handler.close()

        # Remove the file and stream handlers (optional if they need cleanup)
        for handler in self.logger.handlers[:]:
            if isinstance(handler, (logging.FileHandler, logging.StreamHandler)):
                self.logger.removeHandler(handler)
                handler.close()

        event.Skip()  # Ensure the event propagates to the parent if needed

    ## Destructor for the LogTab class.
    #  Ensures the logger handler is removed when the object is destroyed.
    def __del__(self):
        """ Destructor, ensure the logger handler is removed """
        self.logger.removeHandler(self.handler)


    ## Updates the layout of the log tab.
    #  @details This tab only needs to layout and refresh as it doesn't have dynamic text.
    def refresh_gui_text(self):
        # This tab only needs to layout and refresh as it doesn't have dynamic text.
        self.Layout()
        self.Refresh()


## @class WxTextCtrlHandler
#  @brief Handler for log commands that redirects log messages to a wxPython text control.
#
#  The `WxTextCtrlHandler` class is a custom logging handler that redirects log messages
#  to a wxPython text control, allowing log messages to be displayed in the GUI.
class WxTextCtrlHandler(logging.Handler):
    ## Initializes the log handler.
    #  @param ctrl The wxPython text control to which log messages will be redirected.
    def __init__(self, ctrl):
        logging.Handler.__init__(self)
        self.ctrl = ctrl

    ## Emits a log record to the text control.
    #  @param record The log record to be emitted.
    #  @details Formats the log record and writes it to the text control using CallAfter to ensure thread safety.
    def emit(self, record):
        s = self.format(record) + '\n'
        wx.CallAfter(self.ctrl.WriteText, s)


## @class HelpPopUp
#  @brief A popup window for displaying documentation pages.
#
#  The `HelpPopUp` class creates a wxPython-based window that embeds multiple documentation
#  pages in a tabbed interface using `wx.Notebook` and `wx.html2.WebView`. It supports Markdown (`.md`) files
#  by converting them to HTML before rendering.
class HelpPopUp(wx.Frame):
    ## Initializes the documentation popup.
    #  @param parent The parent wx object.
    #  @param translator The translator object to determine the language.
    #  @param file_dir The directory containing the documentation files.
    def __init__(self, parent, translator, file_dir: Path):
        super(HelpPopUp, self).__init__(parent, title="Help", size=(900, 900))

        ## @var notebook
        #  A wx.Notebook widget containing different documentation tabs.
        notebook = wx.Notebook(self)

        # Get the selected language from the translator
        selected_language = translator.get_selected_language()

        # Determine the correct documentation subdirectory based on the language
        doc_dir = file_dir / "documentation" / selected_language

        ## @var dict_docu
        #  A dictionary mapping tab names to documentation file paths based on the selected language.
        dict_docu = {
            translator.translate('application_cluster_tool_gui'): doc_dir / "application.html",
            translator.translate('basics_cluster_analysis'): doc_dir / "foundation_clustering.html",
            translator.translate('documentation_code'): doc_dir.parent / "code" / "index.html"
        }

        # Create tabs with embedded WebView for each documentation page
        for page_name, file_path in dict_docu.items():
            panel = wx.Panel(notebook)
            sizer = wx.BoxSizer(wx.VERTICAL)

            ## @var html_view
            #  A WebView widget to render the documentation content.
            html_view =wx.html2.WebView.New(panel)

            if file_path.suffix == ".html":
                if file_path.exists():
                    # Load URL correctly for local files
                    wx.CallAfter(html_view.LoadURL, str(file_path.resolve()))
                else:
                    html_view.SetPage(f"<h3>Error: File {file_path.name} not found!</h3>", "")
            else:
                html_view.SetPage("<h3>Error: Unsupported file format.</h3>", "")

            sizer.Add(html_view, 1, wx.EXPAND | wx.ALL, 5)
            panel.SetSizer(sizer)

            # Add the panel as a new tab
            notebook.AddPage(panel, page_name)

        ## @var main_sizer
        #  The main layout container for the popup window.
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(notebook, 1, wx.EXPAND)
        self.SetSizer(main_sizer)
        self.Show()



if __name__ == '__main__':
    # Initialize translator outside the app
    translator = Translator(Path(__file__).parent / "Translations.xlsx", language="en")
    app = wx.App()
    frame = LLTFrame(translator)
    app.MainLoop()
