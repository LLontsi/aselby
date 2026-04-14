from django.urls import path
from . import views

app_name = 'parametrage'
urlpatterns = [
    path('', views.config_exercice, name='config'),
    path('historique/', views.historique_exercices, name='historique'),

    # Fonds de roulement
    path('fonds-roulement/', views.fonds_roulement, name='fonds_roulement'),
    path('fonds-roulement/<int:pk>/supprimer/', views.fonds_roulement_supprimer, name='fonds_roulement_supprimer'),

    # Frais exceptionnels
    path('frais-exceptionnels/', views.frais_exceptionnels, name='frais_exceptionnels'),
    path('frais-exceptionnels/<int:pk>/supprimer/', views.frais_exceptionnels_supprimer, name='frais_exceptionnels_supprimer'),

    # Collation
    path('collation/', views.collation, name='collation'),
    path('collation/<int:pk>/supprimer/', views.collation_supprimer, name='collation_supprimer'),

    # Foyer dépenses
    path('foyer-depenses/', views.foyer_depenses, name='foyer_depenses'),
    path('foyer-depenses/<int:pk>/supprimer/', views.foyer_depenses_supprimer, name='foyer_depenses_supprimer'),

    # Agios bancaires
    path('agios/', views.agios, name='agios'),
    path('agios/<int:pk>/supprimer/', views.agios_supprimer, name='agios_supprimer'),

    # Autres dépenses
    path('autres-depenses/', views.autres_depenses, name='autres_depenses'),
    path('autres-depenses/<int:pk>/supprimer/', views.autres_depenses_supprimer, name='autres_depenses_supprimer'),
]
