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

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.fields["stok_urunu"].required = False

        self.fields["stok_urunu"].queryset = (
            Urun.objects.filter(adet__gt=0)
            .order_by("urun_adi", "seri_no")
        )

        self.fields["stok_urunu"].empty_label = (
            "Manuel giriş yapacağım"
        )

    def clean(self):
        cleaned = super().clean()

        stok = cleaned.get("stok_urunu")

        if stok:
            cleaned["urun_adi"] = stok.urun_adi
            cleaned["seri_no"] = stok.seri_no

        return cleaned


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
            raise forms.ValidationError("Yalnızca .xlsx uzantılı dosya yükleyebilirsiniz.")
        return dosya
