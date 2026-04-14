from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.core.paginator import Paginator
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from .models import Adherent
from .forms import AdherentForm

@bureau_required
def liste_adherents(request):
    config = ConfigExercice.get_exercice_courant()
    qs = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
    q = request.GET.get('q', '')
    if q:
        qs = qs.filter(nom_prenom__icontains=q)
    paginator = Paginator(qs, 20)
    page_obj = paginator.get_page(request.GET.get('page'))
    ctx = {'config_exercice': config, 'page_obj': page_obj, 'adherents': page_obj.object_list, 'q': q}
    return render(request, 'dashboard/adherents/liste.html', ctx)

@bureau_required
def fiche_adherent(request, matricule):
    config = ConfigExercice.get_exercice_courant()
    adherent = get_object_or_404(Adherent, matricule=matricule)
    from apps.fonds.models import MouvementFonds
    from apps.tontines.models import ParticipationTontine
    from apps.prets.models import Pret
    mouvements = MouvementFonds.objects.filter(adherent=adherent, annee=config.annee).order_by('mois')
    participations = ParticipationTontine.objects.filter(adherent=adherent, session__annee=config.annee).select_related('session__niveau')
    prets = Pret.objects.filter(adherent=adherent).order_by('-date_octroi')
    ctx = {
        'config_exercice': config, 'adherent': adherent,
        'mouvements': mouvements, 'participations': participations, 'prets': prets,
        'liste_rouge': getattr(adherent, 'liste_rouge', None),
    }
    return render(request, 'dashboard/adherents/fiche.html', ctx)

@bureau_required
def nouvel_adherent(request):
    config = ConfigExercice.get_exercice_courant()
    form = AdherentForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        adherent = form.save()
        messages.success(request, f"Adhérent {adherent.nom_prenom} créé avec succès.")
        return redirect('adherents:fiche', matricule=adherent.matricule)
    return render(request, 'dashboard/adherents/form.html', {'config_exercice': config, 'form': form})

@bureau_required
def modifier_adherent(request, matricule):
    config = ConfigExercice.get_exercice_courant()
    adherent = get_object_or_404(Adherent, matricule=matricule)
    form = AdherentForm(request.POST or None, instance=adherent)
    if request.method == 'POST' and form.is_valid():
        form.save()
        messages.success(request, "Fiche mise à jour.")
        return redirect('adherents:fiche', matricule=matricule)
    return render(request, 'dashboard/adherents/form.html', {'config_exercice': config, 'form': form, 'adherent': adherent})

@bureau_required
def liste_inactifs(request):
    config = ConfigExercice.get_exercice_courant()
    adherents = Adherent.objects.filter(statut='INACTIF').order_by('nom_prenom')
    return render(request, 'dashboard/adherents/inactifs.html', {'config_exercice': config, 'adherents': adherents})
