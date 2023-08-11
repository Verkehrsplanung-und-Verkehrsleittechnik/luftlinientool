# from VisumOverlay import *
import logging

if __name__ == '__main__':
    from pathlib import Path
    import luftlinientool as llt

    # Parameterübergabe
    path_source = Path(r"S:\VuV-Tools\Fertige Tools\Luftlinientool\Python (Visumintegration) in github")
    file_source = 'ZentraleOrteBW_Bezirke_OZ.ver'

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
        ltt1 = llt.LuftlinienCalculator(Visum, attr_quelle="Quelle", attr_ziel="Ziel", anz_versorger=1, max_entfernung=1)
    else:
        print("nicht implementiert")

    ltt1.calculate_main()
    ltt1.export_matrix(visum=ltt1.visum)
    ltt1.export_net(visum=ltt1.visum, links_additive=False)
    ltt1.export_zones_uda_connections("VFS 1")

    ltt1.delete_unused_nodes()

    del Visum
