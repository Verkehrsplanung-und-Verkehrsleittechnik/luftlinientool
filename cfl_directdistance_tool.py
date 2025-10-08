## @package cfl_directdistance_tool.py
# @brief Contains general methods and the DirectDistanceCalculator class for calculating air-line connections
# considering the centrality of zones

import pandas as pd
import logging
import numpy as np
from scipy.spatial import Delaunay
from pathlib import Path
from math import radians
import win32com.client as com
import webbrowser


# from language_management import Translator  # Added import


# todo update immediately in GUI event after change, no repeated update here

# ====== general, useful functions =====

## Opens a Visum instance if not already open
# Enables simultaneous calling of the file internally and externally in Visum
# @param path (Path/str) to a Visum version file
# @param version Visum version, default 240
# @return Visum instance
def open_visum(path, version=240):
    try:
        # tests if the variable Visum exists
        global Visum
        Visum
    except NameError:
        # if not - open a Visum instance
        logging.info('Initializing Visum instance.')
        Visum = com.Dispatch(f"Visum.Visum.{version}")
        logging.info('Opening version file:' + f'{path}')
        Visum.LoadVersion(path)
        logging.info('Version file successfully loaded.')
    return Visum


## Exports the data of a Visum object type in network file format
# @param[in] object Visum object type (singular), e.g. 'link'
# @param[in] df_object_attributes_to_write Data table of the object. Table contains only attributes that can be
# imported into Visum (especially the necessary attributes)
# @param[in] file Target file, in write or append mode (w or a)
def write_object_to_net(object, df_object_attributes_to_write, file):
    header = ["*", "*"]
    header.insert(1, "* Table: " + object + "s")
    header.append(
        ("$" + object.upper().replace(" ", "") + ":" + ";".join(df_object_attributes_to_write.columns)).upper() + "\n")

    # Header is written
    file.write("\n".join(header))
    # Table is written
    df_object_attributes_to_write.to_csv(file, header=False, sep=";", index=False)


## Checks a matrix for symmetry
# @param[in] matrix Matrix to be tested for symmetry
# @param[in] tol Tolerance for allowed deviation, default 1e-8
# @return True or False
def is_symmetric(matrix, tol=1e-8):
    # Application of the maximum norm for the difference between the matrix and its transpose
    # Norm > 0 -> no symmetry
    return np.linalg.norm(matrix.astype(int) - matrix.T.astype(int), np.inf) < tol


## Calculation of the distance between coordinates (Lat, Lon)
# Implementation of the Haversine formula
# @param[in] x1 x-coordinate of point 1
# @param[in] y1 y-coordinate of point 1
# @param[in] vec_x2 x-coordinate vector of points
# @param[in] vec_y2 y-coordinate vector of points
# @return Vector with distances of all points in the point vector to point 1
def calculate_distance_coordinates_haversine(x1, y1, vec_x2, vec_y2):
    # approximate radius of earth in km
    R = 6373.0

    lat1 = radians(y1)
    lon1 = radians(x1)
    vec_lat2 = np.radians(vec_y2)
    vec_lon2 = np.radians(vec_x2)

    # todo Case distinction for negative coordinates
    diff_lon = vec_lon2 - lon1
    diff_lat = vec_lat2 - lat1

    # Haversine formula
    tmp = np.sin(diff_lat / 2) ** 2 + np.cos(lat1) * np.cos(vec_lat2) * np.sin(diff_lon / 2) ** 2
    distances_km = R * 2 * np.arcsin(np.sqrt(tmp))

    return distances_km


## Calculation of the distance between coordinates (x, y)
# Euclidean distance calculation!
# @param[in] x1 x-coordinate of point 1
# @param[in] y1 y-coordinate of point 1
# @param[in] vec_x2 x-coordinate vector of points
# @param[in] vec_y2 y-coordinate vector of points
# @return Vector with distances of all points in the point vector to point 1
def calculate_eucl_distance_coordinates(x1, y1, vec_x2, vec_y2):
    diff_x = vec_x2 - x1
    diff_y = vec_y2 - y1

    # todo Case distinction for negative coordinates

    distances = np.sqrt(np.square(diff_x) + np.square(diff_y))

    return distances


## Identifies the nearest n points from a given set of points to a single point.
# First, the distances of all points to the single point are calculated.
# Then, the n closest points are filtered, and their indices are returned
# @param[in] x_point x-coordinate of the reference point
# @param[in] y_point y-coordinate of the reference point
# @param[in] array_points Array with the x & y coordinates of the points
# @param[in] formula Distance formula to use ("haversine" or "euclidean")
# @param[in] n Desired number of points
# @return list_indices List of indices of the n nearest points
def get_nearest_points_from_set(x_point, y_point, array_points, formula, n=None):
    # If no selection exists
    if (n is not None) and (n >= len(array_points)):
        # all possible points are returned
        return list(range(0, len(array_points)))

    # Calculate distances
    if formula == "haversine":
        distances = calculate_distance_coordinates_haversine(x1=x_point, y1=y_point, vec_x2=array_points[:, 0],
                                                             vec_y2=array_points[:, 1])
    elif formula == "euclidean":
        distances = calculate_eucl_distance_coordinates(x1=x_point, y1=y_point, vec_x2=array_points[:, 0],
                                                        vec_y2=array_points[:, 1])
    else:
        logging.warning("Distance calculation case is not implemented.d")

    # Index of the n lowest values
    list_indices = np.argpartition(distances, n)[:n]

    return list_indices


## Opens the Readme file
# @param path_scripts The directory path where the README.md file is located. Default: Current working directory.
def show_info(path_scripts: Path = Path.cwd()):
    webbrowser.open(str(path_scripts / "README.md"), new=2)


# ===== Class definition ======
## @class DirectDistanceCalculator
# The class contains attributes and calculation methods to determine the VFS between zones
class DirectDistanceCalculator:

    ## Constructor
    # @param source Filename (str) or Visum instance
    # @param attr_cfl Name of the zone attribute that contains the categorization in OZ,MZ,UZ ... Default: TypeNo
    # @param dict_cfl Dictionary containing the attribute values for the respective VFS
    # @param max_distance Specification of the distance up to which neighbors will be connected
    # @param no_suppliers Specification of how many higher-ranking centers a zone should be connected to
    # @param attr_orig Name of the attribute that indicates whether the zone is considered as a source. Default: None
    # @param attr_dest Name of the attribute that indicates whether the zone is considered as a destination. Default: None
    # @param use_filter Indicates whether only active zones are considered. Can only be used if source = Visum instance
    # @param formula_distance Defines the distance function for determining the supply centers.
    # Note: For triangulation, the air-line connections are determined using the Euclidean distance.
    # Delaunay triangulation only works with a projection of Lat/Lon coordinates.
    # @param path_output Optional possibility to specify a path for file export. Default: None. Then the current folder is used if needed.
    # @param translator Optional Translator instance for multilingual logging.
    def __init__(self, source,
                 attr_cfl: str = "TypeNo",
                 dict_cfl: dict = {"cfl_0": 0, "cfl_1": 1, "cfl_2": 2, "cfl_3": 3, "cfl_4": 4, "cfl_5": 5},
                 max_distance=1,
                 no_suppliers=0,
                 attr_orig=None,
                 attr_dest=None,
                 use_filter: bool = False,
                 formula_distance: str = "euclidean",
                 path_output=None):

        ## Debug mode flag. Enables the execution of intermediate analyses that are not considered in the normal program flow
        self.debug_mode = False

        # Required zone attributes
        # Processing the input parameters

        ## Relevant zone attributes
        self.attr_zones = ["No", "Name", "XCoord", "YCoord"]
        ## Centrality attribute
        self.attr_central_level = attr_cfl
        self.attr_zones.append(self.attr_central_level)
        if attr_orig is not None:
            self.attr_zones.append(attr_orig)
        if attr_dest is not None:
            self.attr_zones.append(attr_dest)

        ## Output directory
        self.path_output = path_output

        ## List of VFS to be processed (using stable keys)
        self.cfl = dict_cfl
        ## Dictionary to hold display names for CFLs for export
        self.cfl_labels = {}

        ## Distance calculation
        self.formula_dist = formula_distance

        ## Specification of the neighborhood degree to which equal-ranking connections should be followed
        # (formerly exchange function)
        self.deg_neighbourhood_cfl = dict()
        ## Specification of how many (higher-ranking) suppliers should be connected
        self.num_suppliers_cfl = dict()

        # Goal: dict with VFS: value
        if isinstance(max_distance, int):
            # Conversion to dict with scalar for each VFS
            self.deg_neighbourhood_cfl = dict(zip(self.cfl.keys(), max_distance * np.ones(len(self.cfl), dtype=int)))
        elif isinstance(max_distance, dict):
            self.deg_neighbourhood_cfl = max_distance
        else:
            raise TypeError("Parameter type not implemented.")

        # Goal: dict with VFS: value
        if isinstance(no_suppliers, int):
            # Conversion to dict with scalar for each VFS
            self.num_suppliers_cfl = dict(
                zip(self.cfl.keys(), no_suppliers * np.ones(len(self.cfl), dtype=int)))
        elif isinstance(no_suppliers, dict):
            self.num_suppliers_cfl = no_suppliers
        else:
            raise TypeError("Parameter type not implemented.")

        # Reading the zone data
        # Important: Index of the table = 0...n
        if not isinstance(source, str):
            ## Visum instance
            self.visum = source
            attr_zones = self.attr_zones
            ## Table with zone data
            self.zones = pd.DataFrame(source.Net.Zones.GetMultipleAttributes(attr_zones, OnlyActive=False),
                                      columns=attr_zones)
            set_active_zones = set(
                np.array(source.Net.Zones.GetMultiAttValues("No", OnlyActive=use_filter), dtype=int)[:, 1])
            self.zones["IsActive"] = self.zones["No"].isin(set_active_zones)

            logging.info("%s Zones loaded", len(self.zones))
        else:
            self.visum = None
            logging.warning("Loading zone data failed, input format is not implemented.")

        if attr_orig is None:
            attr_orig = 'origin'
            self.zones[attr_orig] = 1

        if attr_dest is None:
            attr_dest = 'destination'
            self.zones[attr_dest] = 1

        # Handle attr_dest=attr_origin: Delete column duplicate
        if attr_dest == attr_orig:
            self.zones = self.zones.loc[:, ~self.zones.columns.duplicated()]

        ## Source filter attribute
        self.attr_is_from_zone = attr_orig
        ## Destination filter attribute
        self.attr_is_to_zone = attr_dest

        # Init VFS matrices
        # Dict with matrix per VFS: Number of zones x Number of zones
        self.init_results()

        ## Set language of Visum instance
        self.language = self.visum.GetCurrentLanguage()

        # Init dict export
        ## LookupTable Infrastructure: Node number assigned to the zone
        self.dict_export_zone2node = {}  # Contains the number of nodes that are inserted for the zone to be able to insert links
        ## LookUpTable Infrastructure: Mapping internal link number to Visum link number
        self.dict_export_links_vfs = {}  # Contains the links (FromNode-ToNode)
        ## LookUpTable Infrastructure: Mapping connection function level - Visum link type
        self.dict_export_linktypes = {}

        ## DataFrame with the link data of the air-line connections.
        self.edges = pd.DataFrame()

    ## Translates the adjacency matrices of the desired CFL into an edge list
    # @param list_cfl: List of CFL. If not given, all CFL of the instance are used
    # @return df_edges: DataFrame with all edges and their CFL. Attention: Duplicates are not removed
    def adj_matrix_to_links(self, list_cfl=None):

        if list_cfl is None:
            list_cfl = self.cfl.keys()

        list_df_edges = []
        for cfl in list_cfl:
            if not is_symmetric(self.matrices_cfl[cfl]):
                logging.warning(f'{cfl}: Adjacency matrix is not symmetric.')

            df_edges = pd.DataFrame(self.matrices_cfl[cfl]).stack().reset_index()
            df_edges.columns = ["FromNodeNo", "ToNodeNo", "TypeNo"]

            # Filtere Strecken mit True
            df_edges = df_edges.loc[df_edges["TypeNo"] == True, :]

            # setze Attribut VFS
            df_edges.loc[:, "TypeNo"] = cfl

            list_df_edges.append(df_edges)

        df_edges = pd.concat(list_df_edges)

        # Nur eine Strecke zwischen zwei Knoten
        df_edges = df_edges.groupby(["FromNodeNo", "ToNodeNo"]).agg(TypeNo=("TypeNo", min),
                                                                    ListTypeNo=("TypeNo", list)).reset_index()

        df_edges["FromNodeNo"].replace(self.zones["No"], inplace=True)
        df_edges["ToNodeNo"].replace(self.zones["No"], inplace=True)

        return df_edges

    ## Converts the adjacency matrix into a list of connected zones per zone.
    # @param cfl: str, name of the cfl to be considered
    # @param use_zone_names: bool, if true, the zone names are used
    # @return df_set_zones: DataFrame with list object per zone and a column containing the number
    def adj_matrix_to_set_of_connected_zones(self, cfl, use_zone_names=True):
        # Matrix zu DataFrame
        if use_zone_names:
            # Falls Namen verwendet werden sollen, werden die Zeilen & Spalten benannt
            df = pd.DataFrame(self.matrices_cfl[cfl], index=self.zones["Name"], columns=self.zones["Name"])
        else:
            df = pd.DataFrame(self.matrices_cfl[cfl])

        # Erstellt einen DataFrame, der für jede Zeile der Matrix die Spaltennamen enthält, für die der Eintrag True ist
        df_set_zones = df.mul(df.columns.tolist()).apply(lambda x: set(zone for zone in x if zone), axis=1).to_frame(
            name="set zones")
        # Ermittelt die Länge jeder Liste
        df_set_zones["no zones"] = df_set_zones["set zones"].apply(len)

        return df_set_zones

    ## Calculates which neighbours can be reached within n steps.
    # @param max_steps: maximum distance (steps)
    # @param cfl: cfl to be analysed
    # @return matrix: Adjacency matrix for the reachable neighbours within the max-steps
    def calculate_reachability_max_steps(self, max_steps, cfl):

        matrix = np.linalg.matrix_power(self.matrices_cfl[cfl], max_steps)
        np.fill_diagonal(matrix, 0)

        return matrix

    ## Calculates the adjacency matrix for each stored cfl of the instance.
    # @return No return. The results are saved internally.
    def calculate_main(self):
        # Init Ergebnisse
        logging.info("Starting calculation for all CFLs.")
        self.init_results()
        logging.info("Adjacency matrices have been initialized.")

        # Schleife über alle cfl
        for cfl in self.cfl:
            # Berechne die Werte für die VFS
            self.calculate_cfl(cfl)

        logging.info("The calculation for all CFLs is complete.")

    ## Calculates the connections of a CFL.
    # @param cfl: the connection function level for which connections are determined
    def calculate_cfl(self, cfl):

        # Attributswert der Bezirke für die gewählte VFS
        value_vfs = self.cfl[cfl]

        # Attribute der VFS
        k_neighbour = self.deg_neighbourhood_cfl[cfl]
        num_suppliers_cfl = self.num_suppliers_cfl[cfl]

        # Filtere Bezirksdaten, die die Bedingungen erfüllen
        # Sind Aktiv todo Erweiterung Filterung nach attr_filter
        # TypNr <= VFS
        active_zones = self.zones
        active_zones = active_zones.loc[(active_zones[self.attr_central_level] <= value_vfs)
                                        & (active_zones["IsActive"] > 0),
                       :]

        # Abfangen, falls es Bezirke mit identischen Koordinaten gibt, dann funktioniert DeLauney nicht zuverlässig
        if len(active_zones) > len(active_zones[["XCoord", "YCoord"]].drop_duplicates()):
            duplicate_zones = active_zones[active_zones.duplicated(subset=["XCoord", "YCoord"], keep=False)]
            duplicate_zones_string = ', '.join(
                duplicate_zones["No"].apply(lambda x: str(int(x))) + "/" + duplicate_zones["Name"])
            error_msg = f'Aborted: Zones with identical coordinates found (NUMBER/NAME): {duplicate_zones_string}'
            raise ValueError(error_msg)
        elif len(active_zones) < 3:
            logging.info(f'{cfl}: Too few zones are active.')
        else:
            logging.info(f'{cfl}: Performing Delaunay triangulation for {len(active_zones)} zones')

            if k_neighbour > 0:

                # Delaunay Triangulation
                tri = Delaunay(active_zones[["XCoord", "YCoord"]])
                zone_orig_idx_triangles = active_zones.index.values[tri.simplices]
                logging.info(f'{cfl}: {len(zone_orig_idx_triangles)} triangles were created')

                # Adjazenzmatrix ausfüllen
                # Schleife über Dreiecke
                for p1, p2, p3 in zone_orig_idx_triangles:
                    # die drei Punkte des Dreiecks
                    # folgende Abhängigkeiten sind einzufügen:
                    # p1 - p2, p2 - p1, p1 - p3, p3 - p1, p3 - p2, p2 - p3
                    self.matrices_cfl[cfl][p1, p2] = 1
                    self.matrices_cfl[cfl][p1, p3] = 1
                    self.matrices_cfl[cfl][p2, p1] = 1
                    self.matrices_cfl[cfl][p2, p3] = 1
                    self.matrices_cfl[cfl][p3, p1] = 1
                    self.matrices_cfl[cfl][p3, p2] = 1

            # Nachbarschaften Grad n bestimmen
            if k_neighbour > 1:
                logging.info(f'{cfl}: the neighborhood degree must be calculated')
                adj_k_steps = self.calculate_reachability_max_steps(k_neighbour, cfl)
                self.matrices_cfl[cfl] = adj_k_steps

            # Verbindungen mit Versorgungsfunktion
            if num_suppliers_cfl > 0:
                # Erstelle für jeden Bezirk eine Liste der verbundenen Bezirke
                df_list_zones = self.adj_matrix_to_set_of_connected_zones(cfl, use_zone_names=False)

                # Tabelle der möglichen Versorgungszentren
                provider = self.zones.loc[(self.zones[self.attr_central_level] < self.cfl[cfl])
                                          & (self.zones[self.attr_is_from_zone] > 0), :]
                # Menge der möglichen Versorgungszentren
                set_names_provider = set(provider.index)

                # Bestimme für jeden aktiven Bezirk, ob dieser bereits an ein Versorgungszentrum angeschlossen ist
                df_list_zones = df_list_zones.loc[df_list_zones.index.isin(
                    active_zones.loc[active_zones[self.attr_is_from_zone] > 0, :].index), :]
                df_list_zones["no_provider"] = df_list_zones["set zones"].apply(set_names_provider.intersection).apply(
                    len)
                df_list_zones["provider"] = (df_list_zones.index.isin(set_names_provider)) \
                                            | (df_list_zones["no_provider"] >= num_suppliers_cfl)

                # Für alle Bezirke, die die Bedingung nich erfüllen: Verbinde die nächsten k Versorgungszentren
                for zone in df_list_zones.index[df_list_zones["provider"] < True]:
                    zone_data = self.zones.loc[zone, :]

                    # falls bereits mit einem Versorgungszentrum verbunden -> Lösche das Zentrum aus der Menge der Punkte
                    tmp_set_provider = set_names_provider - df_list_zones.loc[zone, "set zones"]
                    provider_tmp = provider.loc[list(tmp_set_provider), :]

                    # Bestimme die fehlende Anzahl an Versorgungszentren
                    # Auswahlkriterium: nächstgelegen
                    list_idx_provider = get_nearest_points_from_set(x_point=zone_data.loc["XCoord"],
                                                                    y_point=zone_data.loc["YCoord"],
                                                                    n=num_suppliers_cfl - df_list_zones.loc[
                                                                        zone, "no_provider"],
                                                                    array_points=provider_tmp[
                                                                        ["XCoord", "YCoord"]].values,
                                                                    formula=self.formula_dist)
                    self.matrices_cfl[cfl][zone, provider_tmp.index[list_idx_provider]] = 1
                    self.matrices_cfl[cfl][provider_tmp.index[list_idx_provider], zone] = 1
                    # debugbefehl Entfernungen
                    # distances = calculate_distance_coordinates(x1=zone_data.loc["XCoord"], y1=zone_data.loc["YCoord"],
                    #                                            vec_x2=provider_tmp.loc[:, "XCoord"].values,
                    #                                            vec_y2=provider_tmp.loc[:, "YCoord"].values)

            # inaktive origin oder Ziel

            # Aufbau Maske mit aktiven und inaktiven OD Paaren
            # origin und Ziel müssen aktiv sein und die transponierte Matrix davon (Symmetrie)
            # Logik: Filtere OD-Paare mit origin & Ziel aktiv...
            #
            #  origin * Ziel  = Matrix
            # (1 0).T * (1 1) = (1  1
            #                    0  0)
            #
            # und symmetrisiere diese
            # (1  1
            #  1  0)

            # Attribute origin und Ziel
            vector_is_from_zone = self.zones[self.attr_is_from_zone].values
            vector_is_to_zone = self.zones[self.attr_is_to_zone].values
            # über dyadisches Produkt ("outer product") verknüpfen
            # Logik als Maske über existierende Matrix legen
            idx_active = np.outer(vector_is_from_zone, vector_is_to_zone).astype(bool)
            # symmetrisieren der Matrix (Bool Oder-Verknüpfung mit transponierter Matrix)
            # Wo OD-Relation, da DO-Relation
            idx_active_symm = idx_active + idx_active.T

            # Adjazenzmatrix wird mit Maske multipliziert, um die Werte der aktiven Paare zu enthalten
            self.matrices_cfl[cfl] = self.matrices_cfl[cfl] * idx_active_symm.astype(int)

            # Symmetrietest
            if np.sum(self.matrices_cfl[cfl] - self.matrices_cfl[cfl].T) > 0:
                raise ValueError("Error: Matrix is not symmetric")

            # debugzwecke
            if self.debug_mode:
                # zeigt an, mit welchen Bezirken ein Bezirk verbunden ist (=benachbarte Zentren)
                list_zones = self.adj_matrix_to_set_of_connected_zones(cfl)

                # Zeigt das Ergebnis in Visum an
                self.export_net(links_additive=False, list_cfl=[cfl])

                logging.info(f'{cfl}: : The result can be viewed in Visum.')

            logging.info(f'The calculation for {cfl} is completed.')

            df_zones_info = self.adj_matrix_to_set_of_connected_zones(cfl)
            df_zones_info["set zones"] = df_zones_info["set zones"].str.join(",")
            # logging.info('\t' + df_zones_info.to_string().replace('\n', '\n\t'))

    ## Deletes nodes in Visum that do not connect any links.
    # All nodes without links are filtered & the active nodes are deleted.
    # The filter is then reset.
    # @return No return. The visum instance is changed.
    def delete_unused_nodes(self):
        if self.visum is None:
            logging.warning("Delete nodes: No Visum instance is linked.")
            return

        # Lösche Punkte ohne Strecke

        # Filter anpassen
        self.visum.Filters.NodeFilter().Init()
        self.visum.Filters.NodeFilter().AddCondition("OP_NONE", False, "Count:InLinks", "EqualVal", 0)
        self.visum.Filters.NodeFilter().AddCondition("OP_AND", False, "Count:OutLinks", "EqualVal", 0)
        self.visum.Filters.NodeFilter().UseFilter = True

        n = self.visum.Net.Nodes.CountActive

        # Löschen
        self.visum.Net.Nodes.RemoveAll(OnlyActive=True)

        # Filter initialisieren
        self.visum.Filters.NodeFilter().Init()

        logging.info(f'{n}: isolated nodes were deleted.')

    ## Exports the desired adjacency matrices either directly to Visum (if Visum instance is linked)
    # or as .mtx file.
    # Existing matrices are overwritten.
    # @param visum: optional visum instance. Default None
    # @param list_cfl: optional set of CFL. Default: None (all of the object)
    def export_matrix(self, list_cfl=None):
        # Falls Visuminstanz erkannt: erstelle & exportiere Daten direkt in Visum (Für Netze mit <1500 Bezirken über SetValues sonst mithilfe einer mtx-Datei im O-Fromat)
        # Sonst: Speichere .mtx Datei

        if list_cfl is None:
            list_cfl = self.cfl.keys()

        logging.info(f'Starting the export of {len(list_cfl)} matrices.')

        for cfl in list_cfl:
            if cfl not in self.matrices_cfl.keys():
                logging.warning(f'Error: {cfl}  is not in the list of calculated CFLs.')
                continue

            matrix_cfl = self.matrices_cfl[cfl]

            # Get the user-facing, translated label for the matrix name
            cfl_label = self.cfl_labels.get(cfl, cfl)  # Fallback to the key if no label is found

            # Benennung in der Matrix in Visum bzw. Datei
            if self.num_suppliers_cfl[cfl] < 1:
                # Term mit Versorgungsfkt wird weggelassen
                name_matrix = f"RIN_{cfl_label}_n={self.deg_neighbourhood_cfl[cfl]}"
            else:
                # Term mit Versorgungsfkt wird hinzugefügt
                name_matrix = f"RIN_{cfl_label}_n={self.deg_neighbourhood_cfl[cfl]}_v={self.num_suppliers_cfl[cfl]}"

            # Übernehme oder definiere einen Output-Pfad (eventuell nicht benötigt)
            path_mat = self.path_output or Path.cwd() / 'mtx'
            path_mat.mkdir(parents=True, exist_ok=True)
            path_mat_file = path_mat / f"{name_matrix}.mtx"

            # Erstellen der Matrix in Visum, falls notwendig

            # Prüfe ob mtx-Datei geschrieben werden muss
            if (self.visum is None) or (self.visum.Net.Zones.Count > 1500):
                df_mat = pd.DataFrame(matrix_cfl,
                                      columns=self.zones["No"].values.astype(int),
                                      index=self.zones["No"].values.astype(int)
                                      , dtype=int
                                      ).stack().reset_index()

                df_mat.columns = ['origin', 'destination', 'Matrix value']

                # Das O-Format kommt ohne 0 Werte aus, bereite einen entsprechenden DataFrame vor
                df_mat_light = df_mat.loc[df_mat['Matrix value'] != 0]

                with open(path_mat_file, "w", newline='\n') as f:
                    str_header = '''$O
* Universität Stuttgart
*
* Verbindungsfunktionsstufe 5
*
* symmetrische Matrix
*
* Parameter
* Nachbarschaftsniveau Z-Z:
* Nachbarschaftsniveau Z-Z+:
*
* Zeitbereich
0 24
*
* Faktor
*
1.0
*
* VonBezirk NachBezirk Matrix value
'''

                    f.write(str_header)
                    df_mat_light.to_csv(f, header=False, sep=" ", index=False)
                    logging.info(f'Matrix {name_matrix} is saved to this file: {path_mat_file}')

            # Falls eine Instanz existiert, Inhalte in Visum importieren
            if self.visum is not None:

                # Anlegen der Matrizen
                if self.visum.Net.Matrices.Count < 1:
                    # Erstelle Matrix
                    matrix_instance = self.visum.Net.AddMatrix(-1, 2, 3)
                    matrix_instance.SetAttValue("CODE", name_matrix)
                    matrix_instance.SetAttValue("NAME", name_matrix)
                else:
                    # Suche existierende Matrizen mit der Benennung
                    matrix_instances = self.visum.Net.Matrices.ItemsByRef(f'''Matrix([CODE]= "{name_matrix}") ''')
                    if matrix_instances.Count < 1:
                        del matrix_instances
                        # Erstelle Matrix
                        matrix_instance = self.visum.Net.AddMatrix(-1, 2, 3)
                        matrix_instance.SetAttValue("CODE", name_matrix)
                        matrix_instance.SetAttValue("NAME", name_matrix)
                    elif matrix_instances.Count > 1:
                        logging.warning("Matrix code exists multiple times, the first matrix will be overwritten.")
                        matrix_instance = matrix_instances.Iterator.Item
                    else:
                        matrix_instance = matrix_instances.Iterator.Item
                        logging.info("Matrix code exists, content will be overwritten.")

                # Import der Werte
                # Wenn es weniger als 1500 Bezirke gibt kann problemlos mit SetValues gearbeitet werden. Ansonsten muss eine mtx-Datei geschreiben werden
                if self.visum.Net.Zones.Count < 1500:
                    matrix_instance.SetValues(matrix_cfl)
                    logging.info(f'{name_matrix}: was read into Visum.')
                else:
                    matrix_instance.Open(path_mat_file, ReadAdditive=False)
                    logging.info(f'{name_matrix}: was read into Visum.')

            else:
                logging.info(f'Visum is not running. Matrices were exported as files to: {path_mat}')

    ## Creates the infrastructure objects in preparation for exporting the infrastructure in the form of dictionaries for nodes, lines, and line types.
    # Called if an object is not present in the dictionaries during export.
    # Prevents the multiple creation of lines and nodes.
    #  @return No return. The results are stored internally.
    def extract_net(self):

        # Create a node list
        df_nodes = self.zones.copy()
        df_nodes = df_nodes.astype({'No': int, self.attr_central_level: int})

        # Create an assignment zones -> nodes
        if self.visum is None:
            no_node_start = 1
            no_link_start = 1
            no_linktype_start = 1
        else:
            no_link_max = int(self.visum.Net.AttValue(r"Max:Links\No") or 0)
            no_linktype_max = int(self.visum.Net.AttValue(r"Max:LinkTypes\No") or 0)
            no_node_max = int(self.visum.Net.AttValue(r"Max:Nodes\No") or 0)

            no_node_start = no_node_max + 1
            no_link_start = no_link_max + 1
            no_linktype_start = no_linktype_max + 1

        # Assignment of the old numbering to the new
        # dict_no_nodes can be used to convince connections, as it links the old numbers (of zones) with the new numbers (nodes)
        # dict[]
        self.dict_export_zone2node = dict(
            zip(df_nodes["No"].astype(int).drop_duplicates(), range(no_node_start, no_node_start + len(df_nodes) + 1)))

        # add link type to dict dict[Name]=Number
        self.dict_export_linktypes = dict(
            zip(self.cfl.keys(), range(no_linktype_start, no_linktype_start + len(self.cfl.keys()) + 1)))

        # adjacency matrix to link table
        df_edges = self.adj_matrix_to_links()

        # replace cfl with typeno
        df_edges["TypeNo"].replace(self.dict_export_linktypes, inplace=True)

        # translate id to zone number
        df_edges["FromNodeNo"].replace(self.dict_export_zone2node, inplace=True)
        df_edges["ToNodeNo"].replace(self.dict_export_zone2node, inplace=True)

        # Add a number
        # 1. Identification of the outward & return direction
        df_edges["No"] = df_edges[["FromNodeNo", "ToNodeNo"]].min(axis=1).astype(str) + "_" + df_edges[
            ["FromNodeNo", "ToNodeNo"]].max(axis=1).astype(str)

        # 2. numbering
        self.dict_export_links_vfs = dict(
            zip(df_edges["No"].drop_duplicates(), range(no_link_start, no_link_start + int(len(df_edges) / 2) + 1)))

        # Test: there is a number for each link
        if len(self.dict_export_links_vfs) != len(df_edges["No"].drop_duplicates()):
            logging.error("Link numbering does not match the number of links.")

        df_edges.loc[:, "Name"] = df_edges["No"]
        df_edges["No"].replace(self.dict_export_links_vfs, inplace=True)

        self.edges = df_edges

    ## Exports a net file
    # if a visum instance exists, the net file is loaded in Visum
    # @param links_additive: if False, the existing routes in Visum are deleted
    # @param list_cfl: List of CFLs that are to be taken into account. Default: All of the object
    # @param create_connectors: if True, connectors are created for each CFL. Default: False
    # @return: None
    def export_net(self, links_additive=True, list_cfl=None, create_connectors=True):
        if self.path_output is None:
            # if no file path is passed: Use visum file path if instance exists, otherwise use the current path
            if self.visum is not None:
                path_net = Path(self.visum.GetPath(1))
            else:
                path_net = Path.cwd()
        else:
            path_net = self.path_output

        if list_cfl is None:
            list_cfl = self.cfl.keys()

        path_net = path_net / f"{'_'.join(list_cfl)}.net"

        # Check: Extract_net notwendig?
        # Erstelle Streckenliste
        df_edges = self.adj_matrix_to_links(list_cfl)
        set_zones = set(df_edges['FromNodeNo']).union(set(df_edges['ToNodeNo']))

        if (len(set_zones - set(self.dict_export_zone2node.keys())) > 0) | (len(df_edges) > len(self.edges)):
            self.extract_net()

        if len(self.edges) < 1:
            logging.info("No links to export, aborting.")
            return

        df_nodes = self.zones.copy()

        if "TypeNo" not in df_nodes.columns.tolist():
            df_nodes["TypeNo"] = df_nodes[self.attr_central_level]

        # Überarbeiten
        df_nodes = df_nodes.astype({'No': int, 'TypeNo': int})
        df_nodes.loc[:, 'Name'] = 'LLT ' + df_nodes['No'].astype(int).astype(str) + ' ' + df_nodes['Name']
        df_nodes["CODE"] = df_nodes["No"].astype(int)
        df_nodes["No"].replace(self.dict_export_zone2node, inplace=True)
        df_nodes = df_nodes[['No', 'Name', 'XCoord', 'YCoord', 'TypeNo', 'CODE']]

        df_edges = self.edges.loc[self.edges["ListTypeNo"].apply(lambda x: bool(set(x).intersection(list_cfl))), :]

        list_tsys_net = pd.DataFrame(self.visum.Net.TSystems.GetMultipleAttributes(["Code"])).squeeze().values.tolist()
        df_linktypes = pd.DataFrame.from_dict(self.dict_export_linktypes, orient="index").reset_index()
        df_linktypes.columns = ["Name", "No"]
        df_linktypes["TSysSet"] = ",".join(list_tsys_net)
        df_linktypes["Rank"] = df_linktypes["No"]

        if create_connectors:
            # Anbindungen vorbereiten von dict_no_nodes
            df_conn = pd.DataFrame(list(self.dict_export_zone2node.items()), columns=["ZONENO", "NODENO"])
            # Duplicate rows for Directions O/D
            df_conn = pd.concat([df_conn] * 2, ignore_index=True)
            # Sort the DataFrame so
            df_conn.sort_values(by=["ZONENO", "NODENO"], inplace=True)
            # Reset index
            df_conn.reset_index(drop=True, inplace=True)
            # Add DIRECTION column
            df_conn["DIRECTION"] = ["O", "D"] * (len(df_conn) // 2)
            # Add TSYSSET for IV-Sys
            tsys_net = pd.DataFrame(self.visum.Net.TSystems.GetMultipleAttributes(["CODE", "TYPE"]),
                                    columns=["CODE", "TYPE"])
            list_ivtsys_net = tsys_net[tsys_net['TYPE'] != 'PUT']["CODE"].to_list()
            df_conn["TSYSSET"] = ",".join(list_ivtsys_net)

        # Write .net file
        with open(path_net, mode="w", newline="\n", encoding="latin-1") as f:
            header = '''$VISION
* Universität Stuttgart Fakultät 2 Bau+Umweltingenieurwissenschaften Stuttgart
* 08/23/22
* Table: Version block
$VERSION:VERSNR;FILETYPE;LANGUAGE;UNIT
13;Net;ENG;KM

'''

            f.write(header)
            write_object_to_net("Node", df_nodes, f)
            write_object_to_net("Link type", df_linktypes, f)
            write_object_to_net("Link", df_edges[["No", "FromNodeNo", "ToNodeNo", "TypeNo", "Name"]], f)
            if create_connectors:
                # Write table: Connectors to the net file
                write_object_to_net("Connector", df_conn, f)

        # If Visum instance is provided: load the .net file
        if self.visum is not None:
            # Conflict management
            # controller = self.visum.IO.CreateAddNetReadController()

            if links_additive is not True:
                self.visum.Net.Links.RemoveAll(OnlyActive=True)

            self.visum.IO.LoadNet(path_net, ReadAdditive=True)

            if self.visum.Net.Links.Count < len(df_edges):
                logging.warning("Error importing the network file.")

        logging.info(f'The network file of {len(list_cfl)} CFLs has been exported to Visum.')

    ## Exports the connections and the number of connections as zone UDAs to Visum
    #  @param cfl The cfl (connection function level) for which the connections should be exported
    #  @return No return value. The Visum instance is modified.
    def export_zones_uda_connections(self, cfl):
        # Create UDA if not exists

        str_no_conn = f"RIN_No.of_Connections_{cfl}".replace(" ", "")
        str_conn = f"RIN_Connections_{cfl}".replace(" ", "")

        try:
            self.visum.Net.Zones.AddUserDefinedAttribute(str_no_conn,
                                                         str_no_conn,
                                                         str_no_conn, 5)
            self.visum.Net.Zones.AddUserDefinedAttribute(str_conn,
                                                         str_conn,
                                                         str_conn, 5)
        except:
            pass

        # Load connections
        df_zones = self.adj_matrix_to_set_of_connected_zones(cfl).reset_index()
        df_zones["No"] = df_zones.Name.replace(self.zones.set_index("Name")["No"].astype(int).to_dict())
        df_zones.set_index("No", inplace=True)

        # Write the result to Visum
        df_format = pd.DataFrame(self.visum.Net.Zones.GetMultiAttValues("No"), columns=["Idx", "No"]).set_index("No")

        df_format = df_format.join(df_zones)
        df_format[str_conn] = df_format["set zones"].str.join(",")

        self.visum.Net.Zones.SetMultiAttValues(str_no_conn, df_format.loc[:, ["Idx", "no zones"]].values)
        self.visum.Net.Zones.SetMultiAttValues(str_conn, df_format.loc[:, ["Idx", str_conn]].values)

    ## Initializes the adjacency matrices
    #  @return No return value. The results are stored internally.
    def init_results(self):
        dict_cfl = {}
        for cfl in self.cfl:
            dict_cfl[cfl] = np.zeros([len(self.zones), len(self.zones)], dtype=bool)

        ## Dictionary with the resulting adjacency matrices of the connection function levels
        self.matrices_cfl = dict_cfl

    ## Filters the links of the inserted link types in Visum.
    #  @return No return value. The Visum instance is modified.
    def filter_links_cfl(self):
        filter = self.visum.Filters.LinkFilter()
        filter.Init()
        filter.AddCondition("OP_NONE", False, "TypeNo", "ContainedIn",
                            ",".join(str(x) for x in self.dict_export_linktypes.values()))
        filter.UseFilter = True

    ## Filters the zones for which the given attribute is greater than 0.
    #  @param filterFromZones Boolean flag to determine whether to filter source zones (True) or destination zones (False). Default: True
    #  @return No return value. The Visum instance is modified.
    def filter_zones_source_targets(self, filterFromZones: bool = True):
        filter = self.visum.Filters.ZoneFilter()
        filter.Init()
        if filterFromZones:
            if self.attr_is_from_zone is not None:
                filter.AddCondition("OP_NONE", False, self.attr_is_from_zone, "GreaterVal", 0)
        else:
            if self.attr_is_to_zone is not None:
                filter.AddCondition("OP_NONE", False, self.attr_is_to_zone, "GreaterVal", 0)

        filter.UseFilter = True

    ## Deletes the links of the VFS.
    #  @return No return value. The Visum instance is modified.
    def delete_added_links(self):
        # Attention: Does NOT delete link types
        self.filter_links_cfl()
        self.visum.Net.Links.RemoveAll(OnlyActive=True)
        self.visum.Filters.LinkFilter().Init()
