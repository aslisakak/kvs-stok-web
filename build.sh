#!/usr/bin/env bash
set -o errexit
python manage.py collectstatic --noinput
python manage.py migrate
python manage.py load_initial_data_if_empty
