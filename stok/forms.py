from django import forms

from .models import (
    DemoCihaz,
    FirmaGonderi,
    SteriSenseKaydi,
    Urun,
)


class UrunForm(forms.ModelForm):
    class Meta:
        model = Urun

        fields = [
            "urun_adi",
            "seri_no",
            "giris_tarihi",
            "adet",
        ]

        widgets = {
            "urun_adi": forms.TextInput(
                attrs={
                    "placeholder": "Örneğin: UTRED30",
                }
            ),
            "seri_no": forms.TextInput(
                attrs={
                    "placeholder": "Seri numarası",
                }
            ),
            "giris_tarihi": forms.DateInput(
                attrs={
                    "type": "date",
                }
            ),
            "adet": forms.NumberInput(
                attrs={
                    "min": 0,
                }
            ),
        }

    def clean_adet(self):
        adet = self.cleaned_data["adet"]

        if adet < 0:
            raise forms.ValidationError(
                "Adet 0 veya daha büyük olmalıdır."
            )

        return adet


class TopluUrunForm(forms.Form):
    urun_adi = forms.CharField(
        max_length=200,
        label="Ürün Adı",
        widget=forms.TextInput(
            attrs={
                "placeholder": "Örneğin: LogTag TRIX-16",
            }
        ),
    )

    giris_tarihi = forms.DateField(
        label="Giriş Tarihi",
        widget=forms.DateInput(
            attrs={
                "type": "date",
            }
        ),
    )

    seri_numaralari = forms.CharField(
        label="Seri Numaraları",
        widget=forms.Textarea(
            attrs={
                "rows": 12,
                "placeholder": (
                    "Her satıra bir seri numarası yazın:\n"
                    "D001031345F5\n"
                    "D00103134450\n"
                    "D001031347ZT"
                ),
            }
        ),
    )

    def clean_seri_numaralari(self):
        metin = self.cleaned_data["seri_numaralari"]

        seri_numaralari = []
        gorulenler = set()

        for satir in metin.splitlines():
            seri_no = satir.strip()

            if not seri_no:
                continue

            anahtar = seri_no.casefold()

            if anahtar not in gorulenler:
                gorulenler.add(anahtar)
                seri_numaralari.append(seri_no)

        if not seri_numaralari:
            raise forms.ValidationError(
                "En az bir seri numarası girmelisiniz."
            )

        return seri_numaralari


class StokMiktarForm(forms.Form):
    miktar = forms.IntegerField(
        min_value=1,
        label="Miktar",
        widget=forms.NumberInput(
            attrs={
                "min": 1,
            }
        ),
    )

    aciklama = forms.CharField(
        required=False,
        max_length=300,
        label="Açıklama",
        widget=forms.TextInput(
            attrs={
                "placeholder": "Açıklama",
            }
        ),
    )


class FirmaGonderiForm(forms.ModelForm):
    class Meta:
        model = FirmaGonderi

        fields = [
            "firma_adi",
            "stok_urunu",
            "urun_adi",
            "seri_no",
            "gonderim_tarihi",
            "notlar",
        ]

        widgets = {
            "firma_adi": forms.TextInput(),
            "stok_urunu": forms.Select(),
            "urun_adi": forms.TextInput(),
            "seri_no": forms.TextInput(),
            "gonderim_tarihi": forms.DateInput(
                attrs={
                    "type": "date",
                }
            ),
            "notlar": forms.Textarea(
                attrs={
                    "rows": 4,
                }
            ),
        }

    def _init_(self, *args, **kwargs):
        super()._init_(*args, **kwargs)

        self.fields["stok_urunu"].required = False

        self.fields["stok_urunu"].queryset = (
            Urun.objects.filter(adet__gt=0)
            .order_by("urun_adi", "seri_no")
        )

        self.fields["stok_urunu"].empty_label = (
            "Manuel giriş yapacağım"
        )

    def clean(self):
        cleaned_data = super().clean()

        stok_urunu = cleaned_data.get("stok_urunu")
        urun_adi = cleaned_data.get("urun_adi", "").strip()

        if stok_urunu:
            cleaned_data["urun_adi"] = stok_urunu.urun_adi
            cleaned_data["seri_no"] = stok_urunu.seri_no

        elif not urun_adi:
            self.add_error(
                "urun_adi",
                "Stoktan ürün seçin veya ürün adını manuel girin.",
            )

        return cleaned_data


class DemoCihazForm(forms.ModelForm):
    class Meta:
        model = DemoCihaz

        fields = [
            "urun_adi",
            "seri_no",
        ]

        widgets = {
            "urun_adi": forms.TextInput(),
            "seri_no": forms.TextInput(),
        }


class SteriSenseForm(forms.ModelForm):
    class Meta:
        model = SteriSenseKaydi

        fields = [
            "bayi",
            "firma",
            "urun_adi",
            "seri_no",
            "tarih",
            "notlar",
        ]

        widgets = {
            "bayi": forms.TextInput(
                attrs={
                    "placeholder": "Bayi",
                }
            ),
            "firma": forms.TextInput(
                attrs={
                    "placeholder": "Firma",
                }
            ),
            "urun_adi": forms.TextInput(
                attrs={
                    "placeholder": "Ürün Adı",
                }
            ),
            "seri_no": forms.TextInput(
                attrs={
                    "placeholder": "Seri No",
                }
            ),
            "tarih": forms.DateInput(
                attrs={
                    "type": "date",
                }
            ),
            "notlar": forms.Textarea(
                attrs={
                    "rows": 4,
                }
            ),
        }


class ExcelIceriAktarForm(forms.Form):
    excel_dosyasi = forms.FileField(
        label="Excel Dosyası",
        help_text="KVS aktarım dosyasını (.xlsx) seçin.",
    )

    def clean_excel_dosyasi(self):
        dosya = self.cleaned_data["excel_dosyasi"]

        if not dosya.name.lower().endswith(".xlsx"):
            raise forms.ValidationError(
                "Yalnızca .xlsx uzantılı dosya yükleyebilirsiniz."
            )

        return dosya