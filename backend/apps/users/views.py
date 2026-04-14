from django.contrib.auth import login, logout, update_session_auth_hash
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import PasswordChangeForm
from django.shortcuts import render, redirect
from django.contrib import messages
from django.utils import timezone
from django.db.models import Sum
from decimal import Decimal

from apps.parametrage.models import ConfigExercice
from apps.fonds.models import MouvementFonds
from apps.tontines.models import ParticipationTontine, SessionTontine, NiveauTontine
from apps.prets.models import Pret
from apps.public.models import Annonce
from apps.saisie.models import TableauDeBord
from .forms import ConnexionForm, ReinitialisationMotDePasseForm, DemandePretForm
from django.views.decorators.http import require_http_methods
MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']


def _ctx_membre(request):
    return {
        'config_exercice': ConfigExercice.get_exercice_courant(),
        'adherent': request.user.adherent,
    }


# ============================================================
# AUTH
# ============================================================

def page_connexion(request):
    if request.user.is_authenticated:
        return redirect('rapports:dashboard') if request.user.est_bureau else redirect('membre:mon_espace')
    form = ConnexionForm(request, data=request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = form.get_user()
        login(request, user)
        messages.success(request, f"Bienvenue, {user.nom_complet} !")
        return redirect('rapports:dashboard') if user.est_bureau else redirect('membre:mon_espace')
    return render(request, 'users/connexion.html', {'form': form})


def deconnexion(request):
    logout(request)
    messages.info(request, "Vous avez été déconnecté.")
    return redirect('public:accueil')


@login_required
def changer_mot_de_passe(request):
    form = PasswordChangeForm(request.user, request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = form.save()
        update_session_auth_hash(request, user)
        messages.success(request, "Mot de passe modifié avec succès.")
        return redirect('rapports:dashboard') if request.user.est_bureau else redirect('membre:mon_espace')
    return render(request, 'users/changer_mot_de_passe.html', {'form': form})


def reinitialiser_mot_de_passe(request):
    form = ReinitialisationMotDePasseForm(request.POST or None)
    if request.method == 'POST' and form.is_valid():
        user = form.get_user()
        if user:
            user.set_password('aselby_2026')
            user.save()
            messages.success(request, "Mot de passe réinitialisé à 'aselby_2026'. Changez-le après connexion.")
            return redirect('users:connexion')
        else:
            messages.error(request, "Aucun compte trouvé avec ces informations.")
    return render(request, 'users/reinitialisation.html', {'form': form})


# ============================================================
# ESPACE MEMBRE
# ============================================================

@login_required
def mon_espace(request):
    if request.user.est_bureau:
        return redirect('rapports:dashboard')
    adherent = request.user.adherent
    if not adherent:
        messages.error(request, "Compte non lié à un adhérent.")
        return redirect('public:accueil')

    config = ConfigExercice.get_exercice_courant()
    now = timezone.now()
    mois = now.month
    annee = now.year

    # Fonds du mois courant
    fonds_courant = MouvementFonds.objects.filter(
        adherent=adherent, mois=mois, annee=annee
    ).first()

    # Tontines du mois
    sessions_mois = SessionTontine.objects.filter(mois=mois, annee=annee)
    participations_mois = ParticipationTontine.objects.filter(
        adherent=adherent, session__in=sessions_mois
    ).select_related('session__niveau')
    nb_parts_total = participations_mois.aggregate(t=Sum('nombre_parts'))['t'] or 0

    # Prêt en cours
    pret_en_cours = Pret.objects.filter(
        adherent=adherent, statut__in=[Pret.EN_COURS, 'EN_RETARD']
    ).first()

    # Saisie du mois courant (TableauDeBord)
    saisie_mois = TableauDeBord.objects.filter(
        adherent=adherent, mois=mois, annee=annee
    ).first()

    # Nombre de mois saisis cette année
    nb_mois_saisis = TableauDeBord.objects.filter(
        adherent=adherent, annee=annee, config_exercice=config
    ).count()

    # Liste rouge
    liste_rouge = getattr(adherent, 'liste_rouge', None)

    # Annonces
    annonces_recentes = Annonce.objects.filter(
        est_publiee=True
    ).order_by('-date_publication')[:3]

    ctx = _ctx_membre(request)
    ctx.update({
        'fonds_courant': fonds_courant,
        'nb_parts_total': nb_parts_total,
        'participations_mois': participations_mois,
        'pret_en_cours': pret_en_cours,
        'liste_rouge': liste_rouge,
        'saisie_mois': saisie_mois,
        'nb_mois_saisis': nb_mois_saisis,
        'annonces_recentes': annonces_recentes,
        'mois_courant_label': f"{MOIS_FR[mois]} {annee}",
    })
    return render(request, 'membre/mon_espace.html', ctx)


@login_required
def mon_fonds(request):
    if request.user.est_bureau:
        return redirect('rapports:dashboard')
    adherent = request.user.adherent
    config = ConfigExercice.get_exercice_courant()

    # Mouvements fonds (MouvementFonds) — épargne, intérêts, capital
    mouvements = MouvementFonds.objects.filter(
        adherent=adherent, annee=config.annee
    ).order_by('mois')

    dernier_mouvement = mouvements.last()
    total_interets = mouvements.aggregate(s=Sum('interet_attribue'))['s'] or Decimal('0')
    nb_mois_saisis = mouvements.count()

    # Saisies mensuelles (TableauDeBord) — versements, mode, pénalités
    saisies = TableauDeBord.objects.filter(
        adherent=adherent, annee=config.annee, config_exercice=config
    ).order_by('mois')

    ctx = _ctx_membre(request)
    ctx.update({
        'mouvements': mouvements,
        'dernier_mouvement': dernier_mouvement,
        'total_interets': total_interets,
        'nb_mois_saisis': nb_mois_saisis,
        'saisies': saisies,
        'liste_rouge': getattr(adherent, 'liste_rouge', None),
    })
    return render(request, 'membre/mon_fonds.html', ctx)


@login_required
def mes_tontines(request):
    if request.user.est_bureau:
        return redirect('rapports:dashboard')
    adherent = request.user.adherent
    config = ConfigExercice.get_exercice_courant()

    # Toutes les participations de l'exercice
    participations = ParticipationTontine.objects.filter(
        adherent=adherent,
        session__niveau__config_exercice=config
    ).select_related('session__niveau').order_by('session__mois')

    # Grouper par niveau
    niveaux_participation = []
    for niv in NiveauTontine.objects.filter(config_exercice=config).order_by('taux_mensuel'):
        parts = participations.filter(session__niveau=niv)
        if parts.exists():
            niveaux_participation.append({'niveau': niv, 'participations': parts})

    nb_participations = participations.count()
    nb_parts_total    = sum(p.nombre_parts for p in participations)
    nb_lots_obtenus   = participations.filter(a_obtenu_lot_principal=True).count()
    nb_mois_banque    = participations.filter(mode_versement='BANQUE').count()

    ctx = _ctx_membre(request)
    ctx.update({
        'participations': participations,
        'niveaux_participation': niveaux_participation,
        'nb_participations': nb_participations,
        'nb_parts_total': nb_parts_total,
        'nb_lots_obtenus': nb_lots_obtenus,
        'nb_mois_banque': nb_mois_banque,
    })
    return render(request, 'membre/mes_tontines.html', ctx)


@login_required
def mes_prets(request):
    if request.user.est_bureau:
        return redirect('rapports:dashboard')
    ctx = _ctx_membre(request)
    ctx['prets'] = Pret.objects.filter(
        adherent=request.user.adherent,
        config_exercice=ConfigExercice.get_exercice_courant()
    ).prefetch_related('remboursements').order_by('-date_octroi')
    return render(request, 'membre/mes_prets.html', ctx)


@login_required
def demander_pret(request):
    if request.user.est_bureau:
        return redirect('rapports:dashboard')
    adherent = request.user.adherent
    config = ConfigExercice.get_exercice_courant()

    # Vérifier prêt actif
    pret_actif = Pret.objects.filter(
        adherent=adherent, statut__in=[Pret.EN_COURS, 'EN_RETARD']
    ).first()

    if request.method == 'POST' and not pret_actif:
        form = DemandePretForm(request.POST, config=config)
        if form.is_valid():
            pret = form.save(commit=False)
            pret.adherent = adherent
            pret.config_exercice = config
            pret.taux_mensuel = config.taux_interet_pret_mensuel
            pret.est_demande_membre = True
            pret.est_valide_bureau = False
            pret.date_demande = timezone.now()
            pret.statut = Pret.EN_COURS
            pret.save()
            messages.success(request, "Votre demande de prêt a été envoyée au bureau.")
            return redirect('membre:mes_prets')
    else:
        form = DemandePretForm(config=config)

    ctx = _ctx_membre(request)
    ctx.update({
        'form': form,
        'pret_actif': pret_actif,
        'config': config,
    })
    return render(request, 'membre/demander_pret.html', ctx)


@login_required
def ma_situation(request):
    if request.user.est_bureau:
        return redirect('rapports:dashboard')
    adherent = request.user.adherent
    config = ConfigExercice.get_exercice_courant()

    # Fiche de cassation
    try:
        from apps.exercice.models import FicheCassation
        fiche_cassation = FicheCassation.objects.filter(
            adherent=adherent, config_exercice=config
        ).first()
    except Exception:
        fiche_cassation = None

    # Capital total (dernier mouvement)
    dernier_mouvement = MouvementFonds.objects.filter(
        adherent=adherent, annee=config.annee
    ).order_by('mois').last()
    capital_total = dernier_mouvement.capital_compose if dernier_mouvement else Decimal('0')

    # Dette prêt
    pret_actif = Pret.objects.filter(
        adherent=adherent, statut__in=[Pret.EN_COURS, 'EN_RETARD']
    ).first()
    dette_pret = pret_actif.solde_restant if pret_actif else Decimal('0')

    # Stats tontines
    participations = ParticipationTontine.objects.filter(
        adherent=adherent, session__niveau__config_exercice=config
    )
    nb_parts_total = sum(p.nombre_parts for p in participations)

    # Stats saisies (TableauDeBord)
    saisies_annee = TableauDeBord.objects.filter(
        adherent=adherent, annee=config.annee, config_exercice=config
    )
    nb_mois_saisis  = saisies_annee.count()
    nb_mois_banque  = saisies_annee.filter(mode_versement='BANQUE').count()
    nb_mois_especes = saisies_annee.filter(mode_versement='ESPECES').count()
    nb_echecs       = saisies_annee.filter(mode_versement='ECHEC').count()

    ctx = _ctx_membre(request)
    ctx.update({
        'fiche_cassation': fiche_cassation,
        'capital_total': capital_total,
        'dette_pret': dette_pret,
        'nb_parts_total': nb_parts_total,
        'nb_mois_saisis': nb_mois_saisis,
        'nb_mois_banque': nb_mois_banque,
        'nb_mois_especes': nb_mois_especes,
        'nb_echecs': nb_echecs,
        'liste_rouge': getattr(adherent, 'liste_rouge', None),
    })
    return render(request, 'membre/ma_situation.html', ctx)

@login_required
def saisir_versement(request):
    from apps.parametrage.models import ConfigExercice
    from django.utils import timezone

    adherent = getattr(request.user, 'adherent', None)
    if not adherent:
        return redirect('membre:mon_espace')

    config = ConfigExercice.get_exercice_courant()
    now = timezone.now()
    mois = int(request.GET.get('mois', now.month))
    annee = int(request.GET.get('annee', now.year))

    saisie_existante = TableauDeBord.objects.filter(
        adherent=adherent, mois=mois, annee=annee
    ).first()

    if request.method == 'POST':
        from decimal import Decimal
        vb  = Decimal(request.POST.get('versement_banque', '0') or '0')
        ve  = Decimal(request.POST.get('versement_especes', '0') or '0')
        num = request.POST.get('numero_cheque', '').strip()

        if vb == 0 and ve == 0:
            messages.error(request, "Saisissez un montant.")
        elif saisie_existante and saisie_existante.est_valide:
            messages.error(request, "Cette saisie a déjà été validée par le bureau.")
        else:
            obj, created = TableauDeBord.objects.update_or_create(
                adherent=adherent, mois=mois, annee=annee,
                defaults=dict(
                    config_exercice=config,
                    versement_banque=vb,
                    versement_especes=ve,
                    numero_cheque_pret=num,
                    est_valide=False,
                )
            )
            messages.success(request, f"Versement {'soumis' if created else 'mis à jour'} — en attente de validation du bureau.")
            return redirect('membre:mon_espace')

    ctx = _ctx_membre(request)
    ctx.update({'mois': mois, 'annee': annee,
                'saisie_existante': saisie_existante, 'config': config})
    return render(request, 'membre/saisir_versement.html', ctx)