from django.shortcuts import render, get_object_or_404
from django.utils import timezone
from django.db.models import Sum
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from apps.adherents.models import Adherent
from .models import MouvementFonds, ReserveMensuelle
from .services import calculer_interets_mensuels
from decimal import Decimal

@bureau_required
def etat_mensuel(request):
    config = ConfigExercice.get_exercice_courant()
    mois = int(request.GET.get('mois', timezone.now().month))
    annee = int(request.GET.get('annee', timezone.now().year))
    mouvements = MouvementFonds.objects.filter(mois=mois, annee=annee, config_exercice=config).select_related('adherent').order_by('adherent__numero_ordre')
    # Si aucune donnée pour le mois courant, afficher le dernier mois disponible
    if not mouvements.exists() and not request.GET.get('mois'):
        dernier = MouvementFonds.objects.filter(
            config_exercice=config
        ).order_by('-annee', '-mois').first()
        if dernier:
            mois = dernier.mois
            annee = dernier.annee
            mouvements = MouvementFonds.objects.filter(mois=mois, annee=annee, config_exercice=config).select_related('adherent').order_by('adherent__numero_ordre')
    total_fonds = mouvements.aggregate(t=Sum('fonds_definitif'))['t'] or Decimal('0')
    total_interets = mouvements.aggregate(t=Sum('interet_attribue'))['t'] or Decimal('0')
    nb_eligibles = mouvements.filter(base_calcul_interet__gt=0).count()
    ctx = {'config_exercice': config, 'mouvements': mouvements, 'total_fonds': total_fonds,
           'total_interets': total_interets, 'nb_eligibles': nb_eligibles, 'mois': mois, 'annee': annee}
    return render(request, 'dashboard/fonds/etat.html', ctx)

@bureau_required
def detail_adherent(request, matricule):
    config = ConfigExercice.get_exercice_courant()
    adherent = get_object_or_404(Adherent, matricule=matricule)
    mouvements = MouvementFonds.objects.filter(adherent=adherent, annee=config.annee).order_by('mois')
    ctx = {'config_exercice': config, 'adherent': adherent, 'mouvements': mouvements}
    return render(request, 'dashboard/fonds/detail.html', ctx)

@bureau_required
def repartition_interets(request):
    config = ConfigExercice.get_exercice_courant()
    if request.method == 'POST':
        mois = int(request.POST.get('mois'))
        annee = int(request.POST.get('annee'))
        pool = Decimal(request.POST.get('pool_interets', '0'))
        result = calculer_interets_mensuels(mois, annee, pool, config)
        from django.contrib import messages
        messages.success(request, f"Intérêts répartis : {result['total_distribue']:,.2f} FCFA entre {result['nb_eligibles']} adhérents éligibles.")
    return render(request, 'dashboard/fonds/interets.html', {'config_exercice': config})