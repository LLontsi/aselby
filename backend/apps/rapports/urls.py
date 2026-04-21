from django.urls import path
from . import views

app_name = 'rapports'

urlpatterns = [
    # Dashboard existant
    path('', views.dashboard, name='dashboard'),

    # Mouvements
    path('mouvements/',                        views.mouvements_liste,       name='mouvements_liste'),
    path('mouvements/saisie/<str:matricule>/', views.mouvements_saisie,      name='mouvements_saisie'),
    path('mouvements/synthese/',               views.mouvements_synthese,    name='mouvements_synthese'),
    path('mouvements/resume/',                 views.mouvements_resume,      name='mouvements_resume'),
    path('mouvements/telecharger/',            views.telecharger_mouvements, name='telecharger_mouvements'),

    # Historique
    path('historique/',                        views.historique_liste,       name='historique_liste'),
    path('historique/saisie/<str:matricule>/', views.historique_saisie,      name='historique_saisie'),
    path('historique/synthese/',               views.historique_synthese,    name='historique_synthese'),
    path('historique/resume/',                 views.historique_resume,      name='historique_resume'),
    path('historique/telecharger/',            views.telecharger_historique, name='telecharger_historique'),
]