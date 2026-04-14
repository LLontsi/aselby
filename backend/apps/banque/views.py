from django.shortcuts import render
from django.utils import timezone
from django.db.models import Sum
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from apps.saisie.models import TableauDeBord
from .models import HistoriqueBancaire, Cheque
from decimal import Decimal

MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']

@bureau_required
def historique(request):
    config = ConfigExercice.get_exercice_courant()
    mois  = int(request.GET.get('mois',  timezone.now().month))
    annee = int(request.GET.get('annee', timezone.now().year))

    # Essayer HistoriqueBancaire d'abord, sinon construire depuis TableauDeBord
    historiques = HistoriqueBancaire.objects.filter(
        mois=mois, annee=annee, config_exercice=config
    ).select_related('adherent').order_by('adherent__numero_ordre')

    # Si HistoriqueBancaire vide → lire depuis TableauDeBord (source de vérité)
    if not historiques.exists() and not request.GET.get('mois'):
        # Chercher le dernier mois disponible dans TableauDeBord
        dernier = TableauDeBord.objects.filter(
            config_exercice=config
        ).order_by('-annee', '-mois').first()
        if dernier:
            mois  = dernier.mois
            annee = dernier.annee

    # Construire depuis TableauDeBord si HistoriqueBancaire vide
    saisies = TableauDeBord.objects.filter(
        mois=mois, annee=annee, config_exercice=config
    ).select_related('adherent').order_by('adherent__numero_ordre')

    ctx = {
        'config_exercice': config,
        'historiques'    : historiques if historiques.exists() else saisies,
        'saisies'        : saisies,
        'depuis_saisies' : not historiques.exists(),
        'mois'           : mois,
        'annee'          : annee,
        'mois_label'     : MOIS_FR[mois] if 1 <= mois <= 12 else str(mois),
    }
    return render(request, 'dashboard/banque/historique.html', ctx)

@bureau_required
def cheques(request):
    config = ConfigExercice.get_exercice_courant()
    mois  = request.GET.get('mois')   # optionnel
    annee = request.GET.get('annee')  # optionnel

    # Par défaut : tous les chèques de l'exercice courant
    cheques_qs = Cheque.objects.filter(
        config_exercice=config
    ).select_related('adherent').order_by('annee', 'mois', 'adherent__numero_ordre')

    # Filtre mois/annee si demandé via URL
    if mois and annee:
        cheques_qs = cheques_qs.filter(mois=int(mois), annee=int(annee))

    # Regrouper par mois pour affichage
    from itertools import groupby
    from django.db.models import Sum as DSum
    mois_disponibles = (Cheque.objects.filter(config_exercice=config)
        .values('mois', 'annee').order_by('annee', 'mois').distinct())

    ctx = {
        'config_exercice'  : config,
        'cheques'          : cheques_qs,
        'mois_filter'      : int(mois) if mois else None,
        'annee_filter'     : int(annee) if annee else None,
        'mois_disponibles' : mois_disponibles,
        'total_cheques'    : cheques_qs.aggregate(t=DSum('montant'))['t'] or 0,
    }
    return render(request, 'dashboard/banque/cheques.html', ctx)

@bureau_required
def tresorerie(request):
    config = ConfigExercice.get_exercice_courant()

    # Synthèse depuis TableauDeBord (données réelles) par mois
    saisies_all = TableauDeBord.objects.filter(
        config_exercice=config
    ).values('mois', 'annee') \
     .annotate(
         banque=Sum('versement_banque'),
         especes=Sum('versement_especes'),
         en_compte=Sum('versement_banque') + Sum('versement_especes'),
         engagement=Sum('montant_engagement'),
     ).order_by('annee', 'mois')

    synthese_mensuelle = []
    for s in saisies_all:
        en_compte = (s['banque'] or Decimal('0')) + (s['especes'] or Decimal('0'))
        engagement = s['engagement'] or Decimal('0')
        synthese_mensuelle.append({
            'mois'       : s['mois'],
            'annee'      : s['annee'],
            'banque'     : s['banque'] or Decimal('0'),
            'especes'    : s['especes'] or Decimal('0'),
            'en_compte'  : en_compte,
            'engagement' : engagement,
            'a_justifier': max(engagement - en_compte, Decimal('0')),
        })

    total_banque  = sum(l['banque']   for l in synthese_mensuelle)
    total_especes = sum(l['especes']  for l in synthese_mensuelle)
    total_compte  = sum(l['en_compte'] for l in synthese_mensuelle)

    ctx = {
        'config_exercice'  : config,
        'synthese_mensuelle': synthese_mensuelle,
        'total_banque'     : total_banque,
        'total_especes'    : total_especes,
        'total_en_compte'  : total_compte,
    }
    return render(request, 'dashboard/banque/tresorerie.html', ctx)