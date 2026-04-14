from django.shortcuts import render
from django.db.models import Sum
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from .models import ContributionFoyer
from decimal import Decimal

@bureau_required
def etat_foyer(request):
    config = ConfigExercice.get_exercice_courant()
    contributions = ContributionFoyer.objects.filter(
        config_exercice=config
    ).select_related('adherent').order_by('adherent__numero_ordre')

    # Fallback vers exercice précédent si vide
    if not contributions.exists():
        config_prec = ConfigExercice.objects.filter(
            annee__lt=config.annee
        ).order_by('-annee').first()
        if config_prec:
            contributions = ContributionFoyer.objects.filter(
                config_exercice=config_prec
            ).select_related('adherent').order_by('adherent__numero_ordre')

    total_auto = contributions.aggregate(t=Sum('montant_auto'))['t'] or Decimal('0')
    total_dons = contributions.aggregate(t=Sum('don_volontaire'))['t'] or Decimal('0')
    ctx = {
        'config_exercice': config,
        'contributions'  : contributions,
        'total_auto'     : total_auto,
        'total_dons'     : total_dons,
        'total'          : total_auto + total_dons,
    }
    return render(request, 'dashboard/foyer/etat.html', ctx)