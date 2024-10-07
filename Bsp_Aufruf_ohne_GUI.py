## @package Bsp_Aufruf_ohne_GUI.py
# @brief Beispielhafte Anwendung Luftlinientool ohne GUI

# from VisumOverlay import *
import logging

if __name__ == '__main__':
    from pathlib import Path
    import luftlinientool as llt

    # Parameterübergabe
    path_source = Path().cwd() #/ "Version"
    file_source = 'Beispielnetz.ver'

    # Settings Logging
    path_logfile = Path(__file__)
    path_logfile = Path.cwd() / path_logfile.name.replace(".py", ".log")

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    logger_format = logging.Formatter("%(asctime)s %(levelname)s: %(message)s", datefmt="%d.%m.%Y %I:%M:%S %p")
    # Output in Konsole & Logfile
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(logger_format)
    file_handler = logging.FileHandler(path_logfile, mode="w")
    file_handler.setFormatter(logger_format)
    # add handles to logger
    root_logger.addHandler(file_handler)
    root_logger.addHandler(stream_handler)

    source = path_source / file_source

    if source.suffix == ".ver":
        Visum = llt.open_visum(source)
        # ltt1 = llt.LuftlinienCalculator(Visum, attr_quelle="Quelle", attr_ziel="Ziel", anz_versorger=1, max_entfernung=1)
        llt1 = llt.LuftlinienCalculator(Visum, anz_versorger=0, max_entfernung=1,
                                        dict_vfs={"VFS 0": 0, "VFS 1": 1, "VFS 2": 2},
                                        attr_quelle="IstUntersuchungsgebiet",
                                        attr_ziel="AddVal1")
    else:
        print("nicht implementiert")

    llt1.calculate_main()
    llt1.export_matrix()
    llt1.export_net()
    llt1.export_zones_uda_connections("VFS 1")

    llt1.delete_unused_nodes()

    del Visum
