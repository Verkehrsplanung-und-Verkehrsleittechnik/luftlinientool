# from VisumOverlay import *

if __name__ == '__main__':
    from pathlib import Path
    import luftlinientool as llt

    # # Parameterübergabe
    path_source = Path(r"C:\Users\ac128405\Desktop\Software\Luftlinientool\Netz_TUM")
    file_source = 'Bezirke_only.ver'
    source = path_source / file_source

    Visum = llt.open_visum(source, version=240)
    ltt1 = llt.LuftlinienCalculator(Visum,
                                    attr_vfs = "TypeNo",
                                    dict_vfs={"VFS 1": 1},
                                    anz_versorger=0, max_entfernung=1)


    ltt1.calculate_vfs("VFS 1")
    ltt1.export_matrix(visum=ltt1.visum, list_vfs=["VFS 1"])
    ltt1.export_net(visum=ltt1.visum, links_additive=True, list_vfs=["VFS 1"])

    ltt1.delete_unused_nodes()

    print("The End")

