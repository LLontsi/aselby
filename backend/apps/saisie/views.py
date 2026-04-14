from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.utils import timezone
from django.db import transaction
from decimal import Decimal
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from apps.adherents.models import Adherent
from .models import TableauDeBord
from .forms import TableauDeBordForm
from apps.fonds import services as fonds_services


@bureau_required
def formulaire_saisie(request):
    config = ConfigExercice.get_exercice_courant()
    mois   = int(request.GET.get('mois',  timezone.now().month))
    annee  = int(request.GET.get('annee', timezone.now().year))

    if request.method == 'POST':
        mat   = request.POST.get('adherent_matricule', '').strip()
        mois  = int(request.POST.get('mois',  mois))
        annee = int(request.POST.get('annee', annee))
        adherent = get_object_or_404(Adherent, matricule=mat)

        mode_raw = request.POST.get('mode_versement', 'ECHEC')
        mode = mode_raw if mode_raw in ['BANQUE', 'ESPECES', 'ECHEC'] else 'ECHEC'
        vb   = Decimal(request.POST.get('versement_banque',  '0') or '0')
        ve   = Decimal(request.POST.get('versement_especes', '0') or '0')
        num  = request.POST.get('numero_cheque', '').strip()

        # Pénalité espèces — décision manuelle de l'admin
        pen_esp_appliquee = 'penalite_especes_appliquee' in request.POST
        pen_esp_val = config.penalite_especes if pen_esp_appliquee else Decimal('0')

        with transaction.atomic():
            saisie, _ = TableauDeBord.objects.update_or_create(
                adherent=adherent, mois=mois, annee=annee,
                defaults=dict(
                    config_exercice=config,
                    versement_banque=vb,
                    versement_especes=ve,
                    numero_cheque_pret=num,
                    mode_versement=mode,
                    penalite_especes_appliquee=pen_esp_appliquee,
                    penalite_especes_appli=pen_esp_val,
                    est_valide=False,
                )
            )
            saisie.calculer_mode_versement()
            from apps.tontines.models import ParticipationTontine
            parts = ParticipationTontine.objects.filter(
                session__mois=mois, session__annee=annee,
                session__niveau__config_exercice=config,
                adherent=adherent
            ).select_related('session__niveau')
            tontines_du = sum(p.session.niveau.taux_mensuel * p.nombre_parts for p in parts)
            saisie.calculer_penalite_echec(tontines_du)
            saisie.calculer_bonus_malus()
            saisie.save()

        messages.success(request, f"Saisie de {adherent.nom_prenom} — {mois:02d}/{annee} enregistrée.")
        return redirect(f"{request.path}?mois={mois}&annee={annee}")

    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
    saisies   = {
        s.adherent_id: s
        for s in TableauDeBord.objects.filter(mois=mois, annee=annee, config_exercice=config)
    }
    for a in adherents:
        a.saisie_courante = saisies.get(a.matricule)

    ctx = {
        'config_exercice': config,
        'adherents': adherents,
        'saisies': saisies,
        'mois': mois,
        'annee': annee,
        'nb_saisis': len(saisies),
        'nb_total': adherents.count(),
    }
    return render(request, 'dashboard/saisie/formulaire.html', ctx)


@bureau_required
def recapitulatif(request):
    config = ConfigExercice.get_exercice_courant()
    mois   = int(request.GET.get('mois',  timezone.now().month))
    annee  = int(request.GET.get('annee', timezone.now().year))
    saisies = (TableauDeBord.objects
               .filter(mois=mois, annee=annee, config_exercice=config)
               .select_related('adherent')
               .order_by('adherent__numero_ordre'))
    ctx = {'config_exercice': config, 'saisies': saisies, 'mois': mois, 'annee': annee}
    return render(request, 'dashboard/saisie/recapitulatif.html', ctx)


@bureau_required
def valider_saisie(request):
    if request.method == 'POST':
        config = ConfigExercice.get_exercice_courant()
        mois   = int(request.POST.get('mois',  timezone.now().month))
        annee  = int(request.POST.get('annee', timezone.now().year))
        with transaction.atomic():
            TableauDeBord.objects.filter(
                mois=mois, annee=annee, config_exercice=config
            ).update(est_valide=True)
        messages.success(request, f"Saisie de {mois}/{annee} validée.")
    return redirect('saisie:recapitulatif')


@bureau_required
def saisies_membres_en_attente(request):
    """Liste des saisies membres soumises et en attente de validation."""
    from django.db.models import Q
    config  = ConfigExercice.get_exercice_courant()
    saisies = (TableauDeBord.objects
               .filter(config_exercice=config, est_valide=False)
               .filter(Q(versement_banque__gt=0) | Q(versement_especes__gt=0))
               .select_related('adherent')
               .order_by('adherent__numero_ordre', '-mois'))
    ctx = {
        'config_exercice': config,
        'saisies': saisies,
        'nb_en_attente': saisies.count(),
    }
    return render(request, 'dashboard/saisie/membres_en_attente.html', ctx)


@bureau_required
def valider_saisie_membre(request, pk):
    """
    Valide ou rejette une saisie soumise par un membre.
    Si versement espèces, l'admin peut décider d'appliquer ou non la pénalité.
    """
    saisie = get_object_or_404(TableauDeBord, pk=pk)
    config = ConfigExercice.get_exercice_courant()

    if request.method == 'POST':
        action = request.POST.get('action')

        if action == 'valider':
            # Pénalité espèces : décision de l'admin au moment de la validation
            pen_esp_appliquee = 'penalite_especes_appliquee' in request.POST
            pen_esp_val = config.penalite_especes if pen_esp_appliquee else Decimal('0')

            saisie.est_valide = True
            saisie.penalite_especes_appliquee = pen_esp_appliquee
            saisie.penalite_especes_appli     = pen_esp_val

            saisie.calculer_mode_versement()
            from apps.tontines.models import ParticipationTontine
            parts = ParticipationTontine.objects.filter(
                session__mois=saisie.mois, session__annee=saisie.annee,
                session__niveau__config_exercice=config,
                adherent=saisie.adherent
            ).select_related('session__niveau')
            tontines_du = sum(
                p.session.niveau.taux_mensuel * p.nombre_parts for p in parts
            )
            saisie.calculer_penalite_echec(tontines_du)
            saisie.calculer_bonus_malus()
            saisie.save()

            pen_msg = (f" (pénalité {config.penalite_especes:,.0f} F appliquée)"
                       if pen_esp_appliquee else "")
            messages.success(
                request,
                f"Saisie de {saisie.adherent.nom_prenom} — "
                f"{saisie.mois}/{saisie.annee} validée.{pen_msg}"
            )

        elif action == 'rejeter':
            motif = request.POST.get('motif', '').strip()
            saisie.versement_banque           = Decimal('0')
            saisie.versement_especes          = Decimal('0')
            saisie.penalite_especes_appliquee = False
            saisie.penalite_especes_appli     = Decimal('0')
            saisie.save()
            messages.warning(request, f"Saisie rejetée.{' ' + motif if motif else ''}")

    return redirect('saisie:membres_en_attente')