

from django.urls import path
from . import views
app_name = 'banque'
urlpatterns = [
    path('',                         views.historique,                      name='historique'),
    path('cheques/',                 views.cheques,                         name='cheques'),
    path('tresorerie/',              views.tresorerie,                      name='tresorerie'),
    path('tabbhistobque/',           views.tabbhistobque,                   name='tabbhistobque'),
    path('tabbordaidedepenses/',     views.tabbordaidedepenses,             name='tabbordaidedepenses'),
    path('telecharger/',             views.telecharger_tabbhistobque,       name='telecharger'),
    path('telecharger/tabbordaide/', views.telecharger_tabbordaidedepenses, name='telecharger_tabbordaidedepenses'),
]