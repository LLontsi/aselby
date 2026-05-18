

from django.urls import path
from . import views
app_name = 'banque'
urlpatterns = [
     path('releve/', views.releve_bancaire, name='releve'),

    # Vue synthèse de la Trésorerie
    path('tresorerie/', views.tresorerie, name='tresorerie'),

    # Export Excel TABBHISTOBQUE (Historique Banque)
    path('telecharger/', views.telecharger_tabbhistobque, name='telecharger'),
]