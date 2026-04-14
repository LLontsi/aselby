from django.shortcuts import render, get_object_or_404
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from .models import ListeRouge

@bureau_required
def liste_rouge(request):
    config = ConfigExercice.get_exercice_courant()
    dettes = ListeRouge.objects.filter(est_solde=False).select_related('adherent').order_by('-date_entree')
    dettes_soldees = ListeRouge.objects.filter(est_solde=True).select_related('adherent').order_by('-date_solde')[:10]
    ctx = {'config_exercice': config, 'dettes': dettes, 'dettes_soldees': dettes_soldees}
    return render(request, 'dashboard/dettes/liste.html', ctx)

@bureau_required
def fiche_dette(request, pk):
    config = ConfigExercice.get_exercice_courant()
    dette = get_object_or_404(ListeRouge, pk=pk)
    ctx = {'config_exercice': config, 'dette': dette, 'remboursements': dette.remboursements.all().order_by('date')}
    return render(request, 'dashboard/dettes/fiche.html', ctx)
