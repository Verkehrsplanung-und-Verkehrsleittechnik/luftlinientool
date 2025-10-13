## @package example_application_wo_gui.py
# @brief Example application of the air-line tool without GUI

# from VisumOverlay import *
import logging
from pathlib import Path
import cfl_directlinenetwork_tool as dlnt

if __name__ == '__main__':

    # Parameter passing
    path_source = Path().cwd() / "Version"
    file_source = 'Beispielnetz.ver'

    # Settings Logging
    path_logfile = Path(__file__)
    path_logfile = Path.cwd() / path_logfile.name.replace(".py", ".log")

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    logger_format = logging.Formatter("%(asctime)s %(levelname)s: %(message)s", datefmt="%d.%m.%Y %H:%M:%S")
    # Output in console & logfile
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(logger_format)
    file_handler = logging.FileHandler(path_logfile, mode="w")
    file_handler.setFormatter(logger_format)
    # Add handlers to logger
    root_logger.addHandler(file_handler)
    root_logger.addHandler(stream_handler)

    source = path_source / file_source

    if source.suffix == ".ver":
        Visum = dlnt.open_visum(source)
        # dln_calculator = dlnt.DirectLineNetworkCalculator(Visum, attr_orig="Quelle", attr_dest="Ziel", no_suppliers=1, max_distance=1)
        dln_calculator = dlnt.DirectLineNetworkCalculator(Visum, dict_cfl={"VFS 0": 0, "VFS 1": 1, "VFS 2": 2}, max_distance=2,
                                                          no_suppliers=0, attr_orig="IstUntersuchungsgebiet", attr_dest="AddVal1")
    else:
        print("not implemented")

    dln_calculator.calculate_main()
    dln_calculator.export_matrix()
    dln_calculator.export_net()
    dln_calculator.export_zones_uda_connections("VFS 1")

    dln_calculator.delete_unused_nodes()

    del Visum
