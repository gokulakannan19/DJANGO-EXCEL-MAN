from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name="home"),
    # path('upload/', views.upload, name="upload"),
    # path('document/', views.document, name="document"),
    # path('document/upload/', views.upload_document, name="upload_document")
]
