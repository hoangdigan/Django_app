from django.urls import path
from . import views

urlpatterns = [
    path("", views.index, name="enterprise"),     
    path("industry/", views.industry, name="industry"),     
    path("macro/", views.macro, name="macro"),     
    path("ratings/", views.ratings, name="ratings"),     
    path("other/", views.other, name="other"),
    path("download/", views.download_file, name="download")
]
