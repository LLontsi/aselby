from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.utils import timezone
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from .models import Pret

@bureau_required
def liste_prets(request):
    config = ConfigExercice.get_exercice_courant()
    # Les prêts peuvent être liés à un exercice précédent et rester actifs
    # → afficher tous les prêts sans filtre sur config_exercice
    prets = Pret.objects.all().select_related('adherent').order_by('-date_octroi')
    ctx = {
        'config_exercice': config,
        'prets': prets,
        'prets_en_cours': prets.filter(statut=Pret.EN_COURS),
        'prets_en_retard': prets.filter(statut=Pret.EN_COURS, nb_mois_retard__gt=0),
        'prets_soldes': prets.filter(statut=Pret.SOLDE),
    }
    return render(request, 'dashboard/prets/liste.html', ctx)

@bureau_required
def detail_pret(request, pk):
    config = ConfigExercice.get_exercice_courant()
    pret = get_object_or_404(Pret, pk=pk)
    ctx = {'config_exercice': config, 'pret': pret, 'remboursements': pret.remboursements.all().order_by('annee','mois')}
    return render(request, 'dashboard/prets/detail.html', ctx)

@bureau_required
def nouveau_pret(request):
    config = ConfigExercice.get_exercice_courant()
    return render(request, 'dashboard/prets/nouveau.html', {'config_exercice': config})

@bureau_required
def valider_demande(request, pk):
    pret = get_object_or_404(Pret, pk=pk, est_demande_membre=True, est_valide_bureau=False)
    if request.method == 'POST':
        pret.est_valide_bureau = True
        pret.date_validation = timezone.now()
        pret.save()
        messages.success(request, f"Demande de prêt de {pret.adherent.nom_prenom} validée.")
    return redirect('rapports:dashboard')

@bureau_required
def rejeter_demande(request, pk):
    pret = get_object_or_404(Pret, pk=pk, est_demande_membre=True, est_valide_bureau=False)
    if request.method == 'POST':
        pret.delete()
        messages.info(request, "Demande de prêt rejetée.")
    return redirect('rapports:dashboard')

@bureau_required
def rapport_retards(request):
    config = ConfigExercice.get_exercice_courant()
    prets_retard = Pret.objects.filter(statut=Pret.EN_COURS, nb_mois_retard__gt=0).select_related('adherent').order_by('-nb_mois_retard')
    ctx = {'config_exercice': config, 'prets_retard': prets_retard}
    return render(request, 'dashboard/prets/rapport_retards.html', ctx)