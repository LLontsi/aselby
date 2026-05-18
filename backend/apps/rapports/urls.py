from django.urls import path
from . import views
app_name = 'rapports'
urlpatterns = [
    path('',                                  views.dashboard,                     name='dashboard'),
    # Visualisation données
    path('tabbord/',                          views.tabbord,                       name='tabbord'),
    path('tabbord/resume/',                   views.tabbord_resume,                name='tabbord_resume'),
    path('autremvt/',                         views.autremvt,                      name='autremvt'),
    path('autremvt/resume/',                  views.autremvt_resume,               name='autremvt_resume'),
    # Exports Excel
    path('tabbord/telecharger/',              views.telecharger_tabbord,           name='telecharger_tabbord'),
    path('autremvt/telecharger/',             views.telecharger_autremouvement,    name='telecharger_autremvt'),
    path('tabbordaidedepenses/telecharger/',  views.telecharger_tabbordaidedepenses, name='telecharger_tabbordaidedepenses'),
    path('listeadherent/telecharger/',        views.telecharger_listeadherent,     name='telecharger_listeadherent'),
    path('travauxfinexercice/telecharger/',   views.telecharger_travauxfinexercice, name='telecharger_travauxfinexercice'),
    # Impression paysage A4
    path('impression/',                       views.impression_index,              name='impression'),
    path('impression/releve/',                views.impression_releve,             name='impression_releve'),
    path('impression/tabbord/',               views.impression_tabbord,            name='impression_tabbord'),
    path('impression/tontines/',              views.impression_tontines,           name='impression_tontines'),
    path('impression/autremvt/',              views.impression_autremvt,           name='impression_autremvt'),
    path('impression/fiches-cassation/',      views.impression_fiches_cassation,   name='impression_fiches_cassation'),
    path('impression/adherents/',             views.impression_adherents,          name='impression_adherents'),
    path('impression/prets/',                 views.impression_prets,              name='impression_prets'),
]