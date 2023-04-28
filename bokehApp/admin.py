from django.contrib import admin
from .models import DataPoint, DataPointTime

admin.site.register(DataPoint)
admin.site.register(DataPointTime)

