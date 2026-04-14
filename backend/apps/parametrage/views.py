from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.db.models import Sum
from apps.core.mixins import bureau_required
from .models import (ConfigExercice, DepenseFondsRoulement,
                     DepenseFraisExceptionnels, DepenseCollation,
                     DepenseFoyer, AgioBancaire, AutreDepense)
from .forms import (ConfigExerciceForm, DepenseFondsRoulementForm,
                    DepenseFraisExceptionnelsForm, DepenseCollationForm,
                    DepenseFoyerForm, AgioBancaireForm, AutreDepenseForm)


@bureau_required
def config_exercice(request):
    config = ConfigExercice.get_exercice_courant()
    fige = bool(config and config.est_ouvert)

    if fige:
        return render(request, 'dashboard/parametrage/config.html',
                      {'config_exercice': config, 'fige': True})

    form = ConfigExerciceForm(request.POST or None, instance=config)
    if request.method == 'POST' and form.is_valid():
        obj = form.save(commit=False)
        if 'ouvrir' in request.POST:
            obj.est_ouvert = True
            from django.utils import timezone
            obj.date_ouverture = timezone.now().date()
        obj.save()
        messages.success(request, "Configuration enregistrée.")
        return redirect('parametrage:config')

    return render(request, 'dashboard/parametrage/config.html',
                  {'config_exercice': config, 'form': form, 'fige': False})


@bureau_required
def historique_exercices(request):
    exercices = ConfigExercice.objects.all().order_by('-annee')
    return render(request, 'dashboard/parametrage/historique.html',
                  {'exercices': exercices})


# ──────────────────────────────────────────────────────────────────────
# DÉPENSES — vue générique + vues spécifiques par rubrique
# ──────────────────────────────────────────────────────────────────────

def _ctx_depenses(config, modele, form_class, titre, rubrique_key):
    """Calcule entrées / sorties / solde pour une rubrique."""
    from apps.adherents.models import Adherent
    nb_adherents = Adherent.objects.filter(statut='ACTIF').count()

    # Entrées = cotisation mensuelle × nb mois saisis × nb adhérents
    from apps.saisie.models import TableauDeBord
    if rubrique_key == 'fonds_roulement':
        montant_mensuel = config.fonds_roulement_mensuel
    elif rubrique_key == 'frais_exceptionnels':
        montant_mensuel = config.frais_exceptionnels_mensuel
    elif rubrique_key == 'collation':
        montant_mensuel = config.collation_mensuelle
    else:
        montant_mensuel = 0

    # Nb mois de saisies validées
    nb_mois = TableauDeBord.objects.filter(
        config_exercice=config, est_valide=True
    ).values('mois').distinct().count()

    total_entrees = montant_mensuel * nb_adherents * nb_mois
    depenses = modele.objects.filter(config_exercice=config).order_by('-date')
    total_sorties = depenses.aggregate(s=Sum('montant'))['s'] or 0
    solde = total_entrees - total_sorties

    return {
        'config_exercice': config,
        'depenses': depenses,
        'total_entrees': total_entrees,
        'total_sorties': total_sorties,
        'solde': solde,
        'montant_mensuel': montant_mensuel,
        'nb_adherents': nb_adherents,
        'nb_mois': nb_mois,
        'titre': titre,
        'rubrique_key': rubrique_key,
    }


@bureau_required
def fonds_roulement(request):
    config = ConfigExercice.get_exercice_courant()
    form = DepenseFondsRoulementForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        obj = form.save(commit=False)
        obj.config_exercice = config
        obj.save()
        messages.success(request, "Dépense enregistrée.")
        return redirect('parametrage:fonds_roulement')
    ctx = _ctx_depenses(config, DepenseFondsRoulement, DepenseFondsRoulementForm,
                        'Fonds de roulement', 'fonds_roulement')
    ctx['form'] = form
    return render(request, 'dashboard/parametrage/depenses.html', ctx)


@bureau_required
def fonds_roulement_supprimer(request, pk):
    dep = get_object_or_404(DepenseFondsRoulement, pk=pk)
    if request.method == 'POST':
        dep.delete()
        messages.success(request, "Dépense supprimée.")
    return redirect('parametrage:fonds_roulement')


@bureau_required
def frais_exceptionnels(request):
    config = ConfigExercice.get_exercice_courant()
    form = DepenseFraisExceptionnelsForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        obj = form.save(commit=False)
        obj.config_exercice = config
        obj.save()
        messages.success(request, "Dépense enregistrée.")
        return redirect('parametrage:frais_exceptionnels')
    ctx = _ctx_depenses(config, DepenseFraisExceptionnels, DepenseFraisExceptionnelsForm,
                        'Frais exceptionnels', 'frais_exceptionnels')
    ctx['form'] = form
    return render(request, 'dashboard/parametrage/depenses.html', ctx)


@bureau_required
def frais_exceptionnels_supprimer(request, pk):
    dep = get_object_or_404(DepenseFraisExceptionnels, pk=pk)
    if request.method == 'POST':
        dep.delete()
        messages.success(request, "Dépense supprimée.")
    return redirect('parametrage:frais_exceptionnels')


@bureau_required
def collation(request):
    config = ConfigExercice.get_exercice_courant()
    form = DepenseCollationForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        obj = form.save(commit=False)
        obj.config_exercice = config
        obj.save()
        messages.success(request, "Dépense enregistrée.")
        return redirect('parametrage:collation')
    ctx = _ctx_depenses(config, DepenseCollation, DepenseCollationForm,
                        'Collation', 'collation')
    ctx['form'] = form
    return render(request, 'dashboard/parametrage/depenses.html', ctx)


@bureau_required
def collation_supprimer(request, pk):
    dep = get_object_or_404(DepenseCollation, pk=pk)
    if request.method == 'POST':
        dep.delete()
        messages.success(request, "Dépense supprimée.")
    return redirect('parametrage:collation')


@bureau_required
def foyer_depenses(request):
    config = ConfigExercice.get_exercice_courant()
    from apps.foyer.models import ContributionFoyer
    total_entrees = ContributionFoyer.objects.filter(
        config_exercice=config
    ).aggregate(s=Sum('montant_auto') + Sum('don_volontaire'))['s'] or 0
    depenses = DepenseFoyer.objects.filter(config_exercice=config).order_by('-date')
    total_sorties = depenses.aggregate(s=Sum('montant'))['s'] or 0

    form = DepenseFoyerForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        obj = form.save(commit=False)
        obj.config_exercice = config
        obj.save()
        messages.success(request, "Virement enregistré.")
        return redirect('parametrage:foyer_depenses')

    return render(request, 'dashboard/parametrage/foyer_depenses.html', {
        'config_exercice': config,
        'depenses': depenses,
        'total_entrees': total_entrees,
        'total_sorties': total_sorties,
        'solde': total_entrees - total_sorties,
        'form': form,
    })


@bureau_required
def foyer_depenses_supprimer(request, pk):
    dep = get_object_or_404(DepenseFoyer, pk=pk)
    if request.method == 'POST':
        dep.delete()
        messages.success(request, "Virement supprimé.")
    return redirect('parametrage:foyer_depenses')


@bureau_required
def agios(request):
    config = ConfigExercice.get_exercice_courant()
    agios_list = AgioBancaire.objects.filter(config_exercice=config).order_by('mois')
    total_agios = agios_list.aggregate(s=Sum('montant_agio'))['s'] or 0
    total_interets = agios_list.aggregate(s=Sum('interet_crediteur'))['s'] or 0

    form = AgioBancaireForm(request.POST or None,
                             initial={'annee': config.annee if config else None})
    if request.method == 'POST' and form.is_valid():
        obj = form.save(commit=False)
        obj.config_exercice = config
        obj.save()
        messages.success(request, "Agio enregistré.")
        return redirect('parametrage:agios')

    return render(request, 'dashboard/parametrage/agios.html', {
        'config_exercice': config,
        'agios': agios_list,
        'total_agios': total_agios,
        'total_interets': total_interets,
        'form': form,
    })


@bureau_required
def agios_supprimer(request, pk):
    agio = get_object_or_404(AgioBancaire, pk=pk)
    if request.method == 'POST':
        agio.delete()
        messages.success(request, "Agio supprimé.")
    return redirect('parametrage:agios')


@bureau_required
def autres_depenses(request):
    config = ConfigExercice.get_exercice_courant()
    dep_list = AutreDepense.objects.filter(config_exercice=config).order_by('-date')
    total = dep_list.aggregate(s=Sum('montant'))['s'] or 0

    form = AutreDepenseForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        obj = form.save(commit=False)
        obj.config_exercice = config
        obj.save()
        messages.success(request, "Dépense enregistrée.")
        return redirect('parametrage:autres_depenses')

    return render(request, 'dashboard/parametrage/autres_depenses.html', {
        'config_exercice': config,
        'depenses': dep_list,
        'total': total,
        'form': form,
    })


@bureau_required
def autres_depenses_supprimer(request, pk):
    dep = get_object_or_404(AutreDepense, pk=pk)
    if request.method == 'POST':
        dep.delete()
        messages.success(request, "Dépense supprimée.")
    return redirect('parametrage:autres_depenses')
