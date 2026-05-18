
from django.urls import path
from . import views
app_name = 'exercice'
urlpatterns = [
    path('cassation/',                      views.fiche_cassation,      name='cassation'),
    path('cassation/saisie/',               views.saisie_cassation,     name='saisie_cassation'),
    path('cassation/saisie/<str:matricule>/',views.saisie_cassation,    name='saisie_cassation_ind'),
    path('synthese/',                       views.synthese_comptes,     name='synthese'),
    path('synthese/saisie/',                views.saisie_synthese,      name='saisie_synthese'),
    path('cloturer/',                       views.cloturer_exercice,    name='cloturer'),
    path('recu/<str:matricule>/',           views.recu_individuel,      name='recu'),
    path('exporter/',                       views.exporter_travaux_fin_exercice, name='exporter'),
    path('etat-versement/<str:matricule>/<int:mois>/<int:annee>/', views.etat_versement_membre, name='etat_versement'),
    path('etat-lot/<str:matricule>/<int:mois>/<int:annee>/', views.etat_lot_membre, name='etat_lot'),
]