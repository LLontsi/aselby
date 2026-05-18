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
from apps.prets.models import Pret
from apps.public.models import Annonce
from apps.saisie.models import SaisieMonthly
from apps.banque.models import ReleveBancaire
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

    # Dernier MouvementFonds disponible (pas forcément le mois courant)
    fonds_courant = (
        MouvementFonds.objects.filter(adherent=adherent, annee=annee)
        .order_by('-mois').first()
    )
    capital_compose = fonds_courant.capital_compose if fonds_courant else None

    # Saisie du mois courant
    saisie_mois = SaisieMonthly.objects.filter(
        adherent=adherent, mois=mois, annee=annee
    ).first()

    # Dernière saisie disponible (pour afficher les lots habituels si mois pas encore saisi)
    derniere_saisie = (
        SaisieMonthly.objects.filter(adherent=adherent, config_exercice=config)
        .order_by('-annee', '-mois').first()
    )

    # Parts tontine totales = cumul annuel
    saisies_annee = SaisieMonthly.objects.filter(
        adherent=adherent, annee=annee, config_exercice=config
    )
    nb_parts_total = sum(
        (s.nbre_lots_t60 or 0) + (s.nbre_lots_t75 or 0) + (s.nbre_lots_t100 or 0)
        for s in saisies_annee
    )

    # Prêt en cours — avec calcul remboursement mensuel
    pret_en_cours = Pret.objects.filter(
        adherent=adherent, statut__in=[Pret.EN_COURS, 'EN_RETARD']
    ).first()
    if pret_en_cours:
        pret_en_cours.remboursement_mensuel = (
            pret_en_cours.solde_restant / Decimal(str(max(1, pret_en_cours.nombre_mois)))
        )
    else:
        # Chercher prêt non soldé dans SaisieMonthly (toutes configs)
        s_pret = SaisieMonthly.objects.filter(
            adherent=adherent, pret_fonds__gt=0
        ).order_by('-annee', '-mois').first()
        if s_pret and float(s_pret.pret_fonds or 0) > float(s_pret.remboursement_pret or 0):
            from datetime import date
            nb_mois = s_pret.nbre_mois_pret or 1
            ech_mois = s_pret.mois + nb_mois
            ech_annee = s_pret.annee + (ech_mois - 1) // 12
            ech_mois = ((ech_mois - 1) % 12) + 1
            solde = float(s_pret.pret_fonds or 0) - float(s_pret.remboursement_pret or 0)
            # Créer un objet simple pour le template
            class PretSimple:
                pass
            pret_en_cours = PretSimple()
            pret_en_cours.montant_principal = float(s_pret.pret_fonds or 0)
            pret_en_cours.solde_restant = solde
            pret_en_cours.date_echeance = date(ech_annee, ech_mois, 17)
            pret_en_cours.statut = 'EN_COURS'
            pret_en_cours.remboursement_mensuel = solde / nb_mois

    # Nombre de mois saisis cette année
    nb_mois_saisis = SaisieMonthly.objects.filter(
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
        'capital_compose': capital_compose,
        'nb_parts_total': nb_parts_total,
        'saisie_mois': saisie_mois,
        'derniere_saisie': derniere_saisie,
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

    # Saisies mensuelles (SaisieMonthly) — versements, mode, pénalités
    saisies = SaisieMonthly.objects.filter(
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
    saisies_tontine = SaisieMonthly.objects.filter(
        adherent=adherent,
        config_exercice=config
    ).order_by('mois')

    # Grouper par niveau tontine (T60, T75, T100)
    niveaux_participation = []
    for code, label in [('T60','Tontine 60 000'), ('T75','Tontine 75 000'), ('T100','Tontine 100 000')]:
        parts = saisies_tontine.filter(**{f'nbre_lots_{code.lower()}__gt': 0})
        if parts.exists():
            niveaux_participation.append({'code': code, 'label': label, 'saisies': parts})

    nb_participations = saisies_tontine.count()
    nb_parts_total    = (
        sum(s.nbre_lots_t60 + s.nbre_lots_t75 + s.nbre_lots_t100 for s in saisies_tontine)
    )
    nb_lots_obtenus   = saisies_tontine.filter(achat_lot_t60__gt=0).count() +                         saisies_tontine.filter(achat_lot_t75__gt=0).count() +                         saisies_tontine.filter(achat_lot_t100__gt=0).count()
    nb_mois_banque    = saisies_tontine.filter(nbre_lots_t60__gt=0).count()

    ctx = _ctx_membre(request)
    ctx.update({
        'saisies_tontine': saisies_tontine,
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
    adherent = request.user.adherent
    config   = ConfigExercice.get_exercice_courant()
    ctx = _ctx_membre(request)

    # Prêts depuis prets_pret (table officielle)
    prets_officiels = Pret.objects.filter(
        adherent=adherent,
        config_exercice=config
    ).order_by('-date_octroi')

    # Si aucun prêt officiel, construire depuis SaisieMonthly (toutes configs)
    if not prets_officiels.exists():
        prets_data = []
        saisies_pret = SaisieMonthly.objects.filter(
            adherent=adherent, pret_fonds__gt=0
        ).order_by('annee', 'mois')
        # Garder seulement les non soldés
        saisies_pret = [s for s in saisies_pret
                        if float(s.pret_fonds or 0) > float(s.remboursement_pret or 0)]
        for s in saisies_pret:
            montant = float(s.pret_fonds or 0)
            remb    = float(s.remboursement_pret or 0)
            nb_mois = s.nbre_mois_pret or 0
            if montant <= 0: continue
            # Solde = pret - remb direct (simplification)
            # En réalité il faudrait sommer tous les remb des mois suivants
            # Pour l'instant: si remb=0 sur la ligne du prêt → solde = pret entier
            solde = max(0, montant - remb)
            # Date échéance
            from datetime import date
            ech_mois  = s.mois + nb_mois
            ech_annee = s.annee + (ech_mois - 1) // 12
            ech_mois  = ((ech_mois - 1) % 12) + 1
            try:
                date_ech = date(ech_annee, ech_mois, 17)
            except Exception:
                date_ech = None
            prets_data.append({
                'montant_principal': montant,
                'montant_total_du':  montant * 1.01 * nb_mois if nb_mois else montant,
                'montant_rembourse': remb,
                'solde_restant':     solde,
                'date_octroi':       date(s.annee, s.mois, 17),
                'date_echeance':     date_ech,
                'nombre_mois':       nb_mois,
                'statut':            'EN_COURS' if solde > 0 else 'REMBOURSE',
                'mode_versement':    s.mode_paiement_pret,
            })
        ctx['prets'] = prets_officiels  # vide
        ctx['prets_data'] = prets_data
    else:
        ctx['prets'] = prets_officiels
        ctx['prets_data'] = []

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
    saisies_tontine = SaisieMonthly.objects.filter(
        adherent=adherent, config_exercice=config
    )
    nb_parts_total = sum(
        (s.nbre_lots_t60 or 0) + (s.nbre_lots_t75 or 0) + (s.nbre_lots_t100 or 0)
        for s in saisies_tontine
    )

    # Stats saisies (SaisieMonthly)
    saisies_annee = SaisieMonthly.objects.filter(
        adherent=adherent, annee=config.annee, config_exercice=config
    )
    nb_mois_saisis  = saisies_annee.count()
    nb_mois_banque  = saisies_annee.filter(nbre_lots_t60__gt=0).count()
    nb_mois_especes = saisies_annee.filter(nbre_lots_t60__gt=0).count()
    nb_echecs       = saisies_annee.filter(nbre_lots_t60=0).count()

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

    saisie_existante = SaisieMonthly.objects.filter(
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
            obj, created = SaisieMonthly.objects.update_or_create(
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

    # Informations tontines et prêt pour affichage dans le formulaire
    # Données tontines du mois depuis SaisieMonthly
    saisies_tontine_mois = SaisieMonthly.objects.filter(
        adherent=adherent, mois=mois, annee=annee
    )
    pret_membre = Pret.objects.filter(
        adherent=adherent, statut__in=[Pret.EN_COURS]
    ).first()
    if pret_membre:
        pret_membre.remboursement_mensuel = (
            pret_membre.solde_restant / Decimal(str(max(1, pret_membre.nombre_mois)))
        )

    ctx = _ctx_membre(request)
    ctx.update({
        'mois': mois, 'annee': annee,
        'saisie_existante': saisie_existante,
        'config': config,
        'saisie_mois': saisie_mois,
        'pret_en_cours': pret_membre,
    })
    return render(request, 'membre/saisir_versement.html', ctx)