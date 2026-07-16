from django.urls import path

from . import views

app_name = "stok"

urlpatterns = [
    # STOKLAR
    path("", views.urun_listesi, name="urun_listesi"),
    path("urun/ekle/", views.urun_ekle, name="urun_ekle"),
    path(
        "urun/<int:urun_id>/duzenle/",
        views.urun_duzenle,
        name="urun_duzenle",
    ),
    path(
        "urun/<int:urun_id>/stok-girisi/",
        views.stok_girisi,
        name="stok_girisi",
    ),
    path(
        "urun/<int:urun_id>/stok-cikisi/",
        views.stok_cikisi,
        name="stok_cikisi",
    ),
    path(
        "urun/<int:urun_id>/hareketler/",
        views.hareketler,
        name="hareketler",
    ),
    path(
        "urun/<int:urun_id>/sil/",
        views.urun_sil,
        name="urun_sil",
    ),
    path(
        "excel/",
        views.excel_aktar,
        name="excel_aktar",
    ),
    path(
        "excel/ice-aktar/",
        views.excel_ice_aktar,
        name="excel_ice_aktar",
    ),

    # FİRMA GÖNDERİLERİ
    path(
        "firma-gonderileri/",
        views.firma_gonderileri,
        name="firma_gonderileri",
    ),
    path(
        "firma-gonderileri/yeni/",
        views.firma_gonderi_ekle,
        name="firma_gonderi_ekle",
    ),
    path(
        "firma-gonderileri/<int:gonderi_id>/duzenle/",
        views.firma_gonderi_duzenle,
        name="firma_gonderi_duzenle",
    ),
    path(
        "firma-gonderileri/<int:gonderi_id>/geri-al/",
        views.firma_gonderi_geri_al,
        name="firma_gonderi_geri_al",
    ),
    path(
        "firma-gonderileri/<int:gonderi_id>/sil/",
        views.firma_gonderi_sil,
        name="firma_gonderi_sil",
    ),
    path(
        "firma-gonderileri/excel/",
        views.firma_gonderileri_excel,
        name="firma_gonderileri_excel",
    ),

    # DEMO CİHAZLAR
    path(
        "demo-cihazlar/",
        views.demo_cihazlar,
        name="demo_cihazlar",
    ),
    path(
        "demo-cihazlar/yeni/",
        views.demo_ekle,
        name="demo_ekle",
    ),
    path(
        "demo-cihazlar/<int:demo_id>/sil/",
        views.demo_sil,
        name="demo_sil",
    ),

    # STERISENSE
path(
    "sterisense/",
    views.sterisense,
    name="sterisense",
),
path(
    "sterisense/yeni/",
    views.sterisense_ekle,
    name="sterisense_ekle",
),
path(
    "sterisense/<int:kayit_id>/duzenle/",
    views.sterisense_duzenle,
    name="sterisense_duzenle",
),
path(
    "sterisense/<int:kayit_id>/sil/",
    views.sterisense_sil,
    name="sterisense_sil",
),
]