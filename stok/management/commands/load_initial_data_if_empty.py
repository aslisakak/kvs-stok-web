from pathlib import Path

from django.contrib.auth import get_user_model
from django.core.management import call_command
from django.core.management.base import BaseCommand

from stok.models import DemoCihaz, FirmaGonderi, SteriSenseKaydi, Urun


class Command(BaseCommand):
    help = "Load initial_data.json once when the database is empty."

    def handle(self, *args, **options):
        User = get_user_model()
        has_data = (
            User.objects.exists()
            or Urun.objects.exists()
            or FirmaGonderi.objects.exists()
            or DemoCihaz.objects.exists()
            or SteriSenseKaydi.objects.exists()
        )
        fixture = Path(__file__).resolve().parents[3] / "initial_data.json"

        if has_data:
            self.stdout.write("Database already contains data; initial import skipped.")
            return
        if not fixture.exists():
            self.stdout.write(self.style.WARNING("initial_data.json not found; import skipped."))
            return

        call_command("loaddata", str(fixture))
        self.stdout.write(self.style.SUCCESS("Initial data imported successfully."))
