from django.urls import path
from . import views

app_name = 'users'

urlpatterns = [
    path('connexion/', views.page_connexion, name='connexion'),
    path('deconnexion/', views.deconnexion, name='deconnexion'),
    path('changer-mot-de-passe/', views.changer_mot_de_passe, name='changer_mot_de_passe'),
    path('mot-de-passe-oublie/', views.reinitialiser_mot_de_passe, name='reinitialisation'),
    path('saisir-versement/', views.saisir_versement, name='saisir_versement'),
]
