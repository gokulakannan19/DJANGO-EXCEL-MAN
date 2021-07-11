from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name="home"),
    path('create_impact/<str:pk>/', views.create_impact, name="create_impact"),
    path('delete_document/<str:pk>/',
         views.delete_document, name="delete_document"),
    # path('upload/', views.upload, name="upload"),
    # path('document/', views.document, name="document"),
    # path('document/upload/', views.upload_document, name="upload_document")delete_document
]
