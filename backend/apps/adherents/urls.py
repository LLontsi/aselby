
from django.urls import path
from . import views
app_name = 'adherents'
urlpatterns = [
    path('',                            views.liste_adherents,            name='liste'),
    path('nouveau/',                    views.nouvel_adherent,            name='nouveau'),
    path('inactifs/',                   views.liste_inactifs,             name='inactifs'),
    path('telecharger/',                views.telecharger_listeadherent,  name='telecharger'),
    path('<str:matricule>/',            views.fiche_adherent,             name='fiche'),
    path('<str:matricule>/modifier/',   views.modifier_adherent,          name='modifier'),
]