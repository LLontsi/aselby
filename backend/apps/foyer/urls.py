from django.urls import path
from . import views
app_name = 'foyer'
urlpatterns = [path('', views.etat_foyer, name='etat')]