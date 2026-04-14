from django.shortcuts import render
from django.db.models import Sum, Count
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from apps.saisie.models import TableauDeBord
from apps.adherents.models import Adherent
from .models import CotisationMutuelle, AideMutuelle
from decimal import Decimal

MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']

@bureau_required
def etat_mutuelle(request):
    config = ConfigExercice.get_exercice_courant()

    # Aides versées — depuis AideMutuelle (config courante ou précédente)
    aides = AideMutuelle.objects.filter(config_exercice=config).select_related('adherent').order_by('-date')
    if not aides.exists():
        config_prec = ConfigExercice.objects.filter(annee__lt=config.annee).order_by('-annee').first()
        if config_prec:
            aides = AideMutuelle.objects.filter(config_exercice=config_prec).select_related('adherent').order_by('-date')

    # Cotisations — construire depuis TableauDeBord.mutuelle (données réelles importées)
    # Grouper par adhérent : total versé en mutuelle sur l'exercice
    saisies = TableauDeBord.objects.filter(
        config_exercice=config, mutuelle__gt=0
    ).values('adherent__matricule', 'adherent__nom_prenom') \
     .annotate(nb=Count('id'), total=Sum('mutuelle')) \
     .order_by('adherent__numero_ordre')

    # Fallback : si aucune saisie avec mutuelle pour config courante, chercher config précédente
    if not saisies.exists():
        config_prec2 = ConfigExercice.objects.filter(annee__lt=config.annee).order_by('-annee').first()
        if config_prec2:
            saisies = TableauDeBord.objects.filter(
                config_exercice=config_prec2, mutuelle__gt=0
            ).values('adherent__matricule', 'adherent__nom_prenom') \
             .annotate(nb=Count('id'), total=Sum('mutuelle')) \
             .order_by('adherent__numero_ordre')

    # Aussi les CotisationMutuelle directes (ex: FAMPOU PAUL inscription)
    cotis_directes = CotisationMutuelle.objects.filter(config_exercice=config)
    if not cotis_directes.exists():
        config_prec3 = ConfigExercice.objects.filter(annee__lt=config.annee).order_by('-annee').first()
        if config_prec3:
            cotis_directes = CotisationMutuelle.objects.filter(config_exercice=config_prec3)

    total_cotisations = (
        (saisies.aggregate(t=Sum('total'))['t'] or Decimal('0')) +
        (cotis_directes.aggregate(t=Sum('montant'))['t'] or Decimal('0'))
    )
    total_aides = aides.aggregate(t=Sum('montant'))['t'] or Decimal('0')

    # Construire cotisations_par_adherent (format attendu par le template)
    cotisations_par_adherent = list(saisies)
    # Ajouter les cotisations directes non déjà dans saisies
    for cd in cotis_directes.select_related('adherent'):
        mat = cd.adherent.matricule
        if not any(c['adherent__matricule'] == mat for c in cotisations_par_adherent):
            cotisations_par_adherent.append({
                'adherent__matricule': mat,
                'adherent__nom_prenom': cd.adherent.nom_prenom,
                'nb': 1,
                'total': cd.montant,
            })

    ctx = {
        'config_exercice'        : config,
        'aides'                  : aides,
        'cotisations_par_adherent': cotisations_par_adherent,
        'total_cotisations'      : total_cotisations,
        'total_aides'            : total_aides,
        'solde_mutuelle'         : total_cotisations - total_aides,
        # Garder aussi l'ancien nom pour compatibilité
        'solde'                  : total_cotisations - total_aides,
    }
    return render(request, 'dashboard/mutuelle/etat.html', ctx)

@bureau_required
def nouvelle_aide(request):
    config = ConfigExercice.get_exercice_courant()
    return render(request, 'dashboard/mutuelle/aide.html', {'config_exercice': config})