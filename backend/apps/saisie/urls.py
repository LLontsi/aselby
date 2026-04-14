from django.urls import path
from . import views
app_name = 'saisie'
urlpatterns = [
    path('', views.formulaire_saisie, name='formulaire'),
    path('recapitulatif/', views.recapitulatif, name='recapitulatif'),
    path('valider/', views.valider_saisie, name='valider'),
    path('membres-en-attente/', views.saisies_membres_en_attente, name='membres_en_attente'),
   path('valider-membre/<int:pk>/', views.valider_saisie_membre, name='valider_saisie_membre'),
]