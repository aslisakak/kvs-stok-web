from django.contrib import admin

from .models import StokHareketi, Urun


@admin.register(Urun)
class UrunAdmin(admin.ModelAdmin):
    list_display = (
        "urun_adi",
        "seri_no",
        "giris_tarihi",
        "adet",
        "stok_durumu",
    )

    search_fields = (
        "urun_adi",
        "seri_no",
    )

    list_filter = (
        "giris_tarihi",
    )

    def stok_durumu(self, urun):
        if urun.adet > 0:
            return "Stokta"
        return "Stokta Yok"

    stok_durumu.short_description = "Durum"


@admin.register(StokHareketi)
class StokHareketiAdmin(admin.ModelAdmin):
    list_display = (
        "urun",
        "islem_turu",
        "miktar",
        "tarih",
        "aciklama",
    )

    search_fields = (
        "urun__urun_adi",
        "urun__seri_no",
        "aciklama",
    )

    list_filter = (
        "islem_turu",
        "tarih",
    )