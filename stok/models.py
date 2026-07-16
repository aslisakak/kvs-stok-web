from django.db import models


class Urun(models.Model):
    urun_adi = models.CharField(
        max_length=200,
        verbose_name="Ürün Adı",
    )

    seri_no = models.CharField(
        max_length=150,
        verbose_name="Seri No",
    )

    giris_tarihi = models.DateField(
        verbose_name="Giriş Tarihi",
    )

    adet = models.PositiveIntegerField(
        default=1,
        verbose_name="Adet",
    )

    olusturulma_tarihi = models.DateTimeField(
        auto_now_add=True,
    )

    def __str__(self):
        return f"{self.urun_adi} - {self.seri_no}"

    @property
    def stokta_mi(self):
        return self.adet > 0

    class Meta:
        verbose_name = "Ürün"
        verbose_name_plural = "Ürünler"
        ordering = ["-id"]


class StokHareketi(models.Model):
    ISLEM_TURLERI = [
        ("giris", "Stok Girişi"),
        ("cikis", "Stok Çıkışı"),
        ("duzenleme", "Ürün Düzenleme"),
        ("firma_gonderisi", "Firma Gönderisi"),
        ("firma_donusu", "Firma Dönüşü"),
    ]

    urun = models.ForeignKey(
        Urun,
        on_delete=models.CASCADE,
        related_name="hareketler",
    )

    islem_turu = models.CharField(
        max_length=30,
        choices=ISLEM_TURLERI,
    )

    miktar = models.IntegerField()

    aciklama = models.CharField(
        max_length=300,
        blank=True,
    )

    tarih = models.DateTimeField(
        auto_now_add=True,
    )

    def __str__(self):
        return f"{self.urun} - {self.get_islem_turu_display()}"

    class Meta:
        verbose_name = "Stok Hareketi"
        verbose_name_plural = "Stok Hareketleri"
        ordering = ["-tarih"]


class FirmaGonderi(models.Model):
    firma_adi = models.CharField(
        max_length=200,
        verbose_name="Firma Adı",
    )

    stok_urunu = models.ForeignKey(
        Urun,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="firma_gonderileri",
        verbose_name="Stoktan Seç",
    )

    urun_adi = models.CharField(
        max_length=200,
        default="",
        verbose_name="Ürün Adı",
    )

    seri_no = models.CharField(
        max_length=150,
        default="",
        blank=True,
        verbose_name="Seri No",
    )

    gonderim_tarihi = models.DateField(
        verbose_name="Gönderim Tarihi",
    )

    notlar = models.TextField(
        blank=True,
        verbose_name="Notlar",
    )

    aktif = models.BooleanField(
        default=True,
        verbose_name="Aktif",
    )

    def __str__(self):
        return f"{self.firma_adi} - {self.urun_adi} - {self.seri_no}"

    class Meta:
        verbose_name = "Firma Gönderisi"
        verbose_name_plural = "Firma Gönderileri"
        ordering = ["-gonderim_tarihi", "-id"]


class DemoCihaz(models.Model):
    urun_adi = models.CharField(
        max_length=200,
        verbose_name="Ürün Adı",
    )

    seri_no = models.CharField(
        max_length=150,
        verbose_name="Seri No",
    )

    def __str__(self):
        return f"{self.urun_adi} - {self.seri_no}"

    class Meta:
        verbose_name = "Demo Cihaz"
        verbose_name_plural = "Demo Cihazlar"
        ordering = ["urun_adi", "seri_no"]


class SteriSenseKaydi(models.Model):
    bayi = models.CharField(
        max_length=200,
        verbose_name="Bayi",
    )

    firma = models.CharField(
        max_length=200,
        verbose_name="Firma",
    )

    urun_adi = models.CharField(
        max_length=200,
        verbose_name="Ürün Adı",
    )

    seri_no = models.CharField(
        max_length=150,
        verbose_name="Seri No",
    )

    tarih = models.DateField(
        verbose_name="Tarih",
    )

    notlar = models.TextField(
        blank=True,
        verbose_name="Notlar",
    )

    def __str__(self):
        return f"{self.bayi} - {self.firma} - {self.urun_adi}"

    class Meta:
        verbose_name = "SteriSense Kaydı"
        verbose_name_plural = "SteriSense Kayıtları"
        ordering = ["-tarih", "-id"]