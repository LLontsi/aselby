from django.urls import path
from . import views
app_name = 'saisie'
urlpatterns = [
    path('',              views.saisie_mensuelle,    name='saisie'),
    path('recap/',        views.recapitulatif,        name='recap'),
    path('interets/',     views.repartition_interets, name='interets'),
    path('attente/',      views.saisies_en_attente,   name='attente'),
    path('valider/',      views.valider_saisie,       name='valider_attente'),
    path('rejeter/',      views.rejeter_saisie,       name='rejeter_attente'),
]