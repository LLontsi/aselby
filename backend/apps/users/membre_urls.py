from django.urls import path
from . import views

app_name = 'membre'

urlpatterns = [
    path('', views.mon_espace, name='mon_espace'),
    path('mes-tontines/', views.mes_tontines, name='mes_tontines'),
    path('mes-prets/', views.mes_prets, name='mes_prets'),
    path('demander-un-pret/', views.demander_pret, name='demander_pret'),
    path('mon-fonds/', views.mon_fonds, name='mon_fonds'),
    path('ma-situation/', views.ma_situation, name='ma_situation'),
    path('saisir-versement/', views.saisir_versement, name='saisir_versement'),
]
