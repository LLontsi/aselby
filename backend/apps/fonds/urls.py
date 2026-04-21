
from django.urls import path
from . import views
app_name = 'fonds'
urlpatterns = [
    path('',                           views.etat_mensuel,                  name='etat'),
    path('interets/',                  views.repartition_interets,          name='interets'),
    path('telecharger/fondscaisse/',   views.telecharger_listefondscaisse,  name='telecharger_fondscaisse'),
    path('telecharger/basecalcul/',    views.telecharger_basecalculinteret, name='telecharger_basecalcul'),
    path('<str:matricule>/',           views.detail_adherent,               name='detail'),
]