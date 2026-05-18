"""
apps/saisie/views.py
Formulaire 2 : Saisie mensuelle principale (TABBORD).
À saisir APRÈS le relevé bancaire.
Grille : tous les adhérents, tous les champs TABBORD manuels.
"""
from django.shortcuts import render, redirect
from django.contrib import messages
from django.utils import timezone
from django.db import transaction
from decimal import Decimal
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from apps.adherents.models import Adherent
from .models import SaisieMonthly
from .services import calculer_mouvement_fonds, recalculer_tout_le_mois

MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']
D = Decimal


def _d(post, key):
    """Extrait un Decimal du POST, 0 si absent ou invalide."""
    v = post.get(key, '0') or '0'
    try:
        return D(str(v).replace(' ', '').replace('\xa0', ''))
    except Exception:
        return D('0')

def _i(post, key):
    """Extrait un entier du POST."""
    try:
        return int(post.get(key, 0) or 0)
    except Exception:
        return 0

def _s(post, key):
    """Extrait une chaîne du POST."""
    return post.get(key, '').strip()


@bureau_required
def saisie_mensuelle(request):
    """
    Grille principale : tous les adhérents × tous les champs TABBORD manuels.
    Sauvegarde en une seule soumission + calcul auto MouvementFonds.
    """
    config    = ConfigExercice.get_exercice_courant()
    mois      = int(request.GET.get('mois',  timezone.now().month))
    annee     = int(request.GET.get('annee', config.annee))
    adherents = sorted(Adherent.objects.filter(statut='ACTIF'), key=lambda a: (0 if a.matricule=='AS201648' else 1, a.numero_ordre))

    if request.method == 'POST':
        mois  = int(request.POST.get('mois',  mois))
        annee = int(request.POST.get('annee', annee))

        nb_saisis = 0
        with transaction.atomic():
            for adh in adherents:
                pfx = f"{adh.matricule}_"

                # Vérifier qu'au moins nbre_lots_t60 est saisi (sinon ignorer)
                nbre_t60  = _i(request.POST, f"{pfx}nbre_lots_t60")
                nbre_t75  = _i(request.POST, f"{pfx}nbre_lots_t75")
                nbre_t100 = _i(request.POST, f"{pfx}nbre_lots_t100")

                if nbre_t60 == 0 and nbre_t75 == 0 and nbre_t100 == 0:
                    # Adhérent ECHEC = aucune part ce mois
                    # On crée quand même l'enregistrement pour tracer
                    pass

                saisie, _ = SaisieMonthly.objects.update_or_create(
                    adherent=adh, mois=mois, annee=annee,
                    defaults={
                        'config_exercice': config,
                        # Tontines
                        'nbre_lots_t60':  nbre_t60,
                        'nbre_lots_t75':  nbre_t75,
                        'nbre_lots_t100': nbre_t100,
                        # Lots
                        'achat_lot_t60':        _d(request.POST, f"{pfx}achat_lot_t60"),
                        'achat_lot_t75':        _d(request.POST, f"{pfx}achat_lot_t75"),
                        'achat_lot_t100':       _d(request.POST, f"{pfx}achat_lot_t100"),
                        'mode_paiement_lot':    _s(request.POST, f"{pfx}mode_paiement_lot"),
                        'vente_petit_lot_t60':  _d(request.POST, f"{pfx}vente_petit_lot_t60"),
                        'vente_petit_lot_t75':  _d(request.POST, f"{pfx}vente_petit_lot_t75"),
                        'vente_petit_lot_t100': _d(request.POST, f"{pfx}vente_petit_lot_t100"),
                        'interet_petit_lot_t60':  _d(request.POST, f"{pfx}interet_petit_t60"),
                        'interet_petit_lot_t75':  _d(request.POST, f"{pfx}interet_petit_t75"),
                        'interet_petit_lot_t100': _d(request.POST, f"{pfx}interet_petit_t100"),
                        'mode_remb_petit_lot':  _s(request.POST, f"{pfx}mode_remb_petit_lot"),
                        # Prêts
                        'remboursement_pret':   _d(request.POST, f"{pfx}remboursement_pret"),
                        'mode_remb_pret':       _s(request.POST, f"{pfx}mode_remb_pret"),
                        'pret_fonds':           _d(request.POST, f"{pfx}pret_fonds"),
                        'mode_paiement_pret':   _s(request.POST, f"{pfx}mode_paiement_pret"),
                        'nbre_mois_pret':       _i(request.POST, f"{pfx}nbre_mois_pret"),
                        # Versements spéciaux
                        'complement_epargne':   _d(request.POST, f"{pfx}complement_epargne"),
                        'montant_especes':      _d(request.POST, f"{pfx}montant_especes"),
                        'numero_cheque':        _s(request.POST, f"{pfx}numero_cheque"),
                        # Dépenses
                        'libelle_depense':      _s(request.POST, f"{pfx}libelle_depense"),
                        'compte_depense':       _s(request.POST, f"{pfx}compte_depense"),
                        'montant_depense':      _d(request.POST, f"{pfx}montant_depense"),
                        'depense_collation':    _d(request.POST, f"{pfx}depense_collation"),
                        # Cas exceptionnels
                        'sanction':             _d(request.POST, f"{pfx}sanction"),
                        'inscription':          _d(request.POST, f"{pfx}inscription"),
                        'retrait_partiel':      _d(request.POST, f"{pfx}retrait_partiel"),
                        'mutuelle':             _d(request.POST, f"{pfx}mutuelle"),
                        'contribution_foyer':   _d(request.POST, f"{pfx}contribution_foyer"),
                        'penalite_pret_fonds':  _d(request.POST, f"{pfx}penalite_pret_fonds"),
                        # Champs saisie manuelle (nouveaux)
                        'penalite_versement_especes_saisi': _d(request.POST, f"{pfx}penalite_versement_especes_saisi"),
                        'interet_pret_saisi':    _d(request.POST, f"{pfx}interet_pret_saisi"),
                        'bonus_malus_saisi':     _d(request.POST, f"{pfx}bonus_malus_saisi"),
                        # montant_lot_t75/t100 alimentés via formulaire TONTINE
                    }
                )

                # Calcul MouvementFonds automatiquement
                calculer_mouvement_fonds(saisie)

                # Enregistrer remboursement prêt si saisi
                if saisie.remboursement_pret > 0:
                    _enregistrer_remboursement_pret(
                        adh, saisie.remboursement_pret,
                        mois, annee, saisie.mode_remb_pret, config)

                nb_saisis += 1

        messages.success(
            request,
            f"Saisie mensuelle {MOIS_FR[mois]} {annee} enregistrée "
            f"({nb_saisis} adhérent(s)). MouvementFonds calculés automatiquement."
        )
        return redirect(f"{request.path}?mois={mois}&annee={annee}")

    # GET — charger données existantes
    saisies = {
        s.adherent_id: s
        for s in SaisieMonthly.objects.filter(
            mois=mois, annee=annee, config_exercice=config)
    }
    from apps.banque.models import ReleveBancaire
    releves = {
        r.adherent_id: r
        for r in ReleveBancaire.objects.filter(
            mois=mois, annee=annee, config_exercice=config)
    }

    prec = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv = (1,  annee+1) if mois == 12 else (mois+1, annee)

    return render(request, 'dashboard/saisie/saisie_mensuelle.html', {
        'config_exercice': config,
        'adherents':       adherents,
        'saisies':         saisies,
        'releves':         releves,
        'mois':            mois,
        'annee':           annee,
        'mois_label':      MOIS_FR[mois],
        'mois_fr':         MOIS_FR,
        'prec_mois':       prec[0], 'prec_annee': prec[1],
        'suiv_mois':       suiv[0], 'suiv_annee': suiv[1],
        'nb_total':        len(adherents),
        'nb_saisis':       len(saisies),
        'libelles_depense': ['COMMUNICATION', 'TAXI BANQUE', 'COLLATION', 'CREDIT TELEPHONE', 'ACHAT FOURNITURE', 'BOISSON', 'COMMISSION', 'DECES', 'OUVERTURE COMPTE', 'PRODUCTION RAPPORT', 'TERRAIN'],
        'nb_releves':      len(releves),
    })


def _enregistrer_remboursement_pret(adh, montant, mois, annee, mode, config):
    """Enregistre le remboursement prêt et met à jour le solde."""
    from apps.prets.models import Pret
    pret = Pret.objects.filter(
        adherent=adh, config_exercice=config,
        statut=Pret.EN_COURS
    ).first()
    if pret and montant > 0:
        pret.enregistrer_remboursement(montant, mois, annee, mode)


@bureau_required
def recapitulatif(request):
    """Synthèse du mois — toutes les saisies avec les calculs auto."""
    config    = ConfigExercice.get_exercice_courant()
    mois      = int(request.GET.get('mois',  timezone.now().month))
    annee     = int(request.GET.get('annee', config.annee))

    saisies = SaisieMonthly.objects.filter(
        mois=mois, annee=annee, config_exercice=config
    ).select_related('adherent').order_by('adherent__numero_ordre')

    from apps.fonds.models import MouvementFonds
    from django.db.models import Sum
    mvts = MouvementFonds.objects.filter(
        mois=mois, annee=annee, config_exercice=config)
    totaux_mvt = mvts.aggregate(
        total_fonds    = Sum('fonds_definitif'),
        total_capital  = Sum('capital_compose'),
        total_interet  = Sum('interet_attribue'),
        nb_eligibles   = Sum('base_calcul_interet'),
    )

    adherents  = Adherent.objects.filter(statut='ACTIF')
    nb_manquants = adherents.count() - saisies.count()

    prec = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv = (1,  annee+1) if mois == 12 else (mois+1, annee)

    return render(request, 'dashboard/saisie/recapitulatif.html', {
        'config_exercice': config,
        'saisies':         saisies,
        'totaux_mvt':      totaux_mvt,
        'mois':            mois,
        'annee':           annee,
        'mois_label':      MOIS_FR[mois],
        'mois_fr':         MOIS_FR,
        'prec_mois':       prec[0], 'prec_annee': prec[1],
        'suiv_mois':       suiv[0], 'suiv_annee': suiv[1],
        'nb_total':        len(adherents),
        'nb_saisis':       saisies.count(),
        'nb_manquants':    nb_manquants,
    })


@bureau_required
def repartition_interets(request):
    """Répartition du pool d'intérêts mensuel."""
    config = ConfigExercice.get_exercice_courant()
    if request.method == 'POST':
        from .services import repartir_interets
        mois  = int(request.POST.get('mois'))
        annee = int(request.POST.get('annee'))
        pool  = D(request.POST.get('pool_interets', '0') or '0')
        result = repartir_interets(mois, annee, pool, config)
        messages.success(
            request,
            f"Intérêts {MOIS_FR[mois]} {annee} répartis : "
            f"{result.pool_interets:,.0f} F entre "
            f"{result.nb_adherents_eligibles} adhérents."
        )
        return redirect(request.path)

    from apps.fonds.models import MouvementFonds, ReserveMensuelle
    mois  = int(request.GET.get('mois',  timezone.now().month))
    annee = int(request.GET.get('annee', config.annee))
    reserve = ReserveMensuelle.objects.filter(
        mois=mois, annee=annee, config_exercice=config).first()
    mvts = MouvementFonds.objects.filter(
        mois=mois, annee=annee, config_exercice=config,
        base_calcul_interet__gt=0
    ).select_related('adherent').order_by('adherent__numero_ordre')

    return render(request, 'dashboard/saisie/interets.html', {
        'config_exercice': config,
        'reserve':   reserve,
        'mvts':      mvts,
        'mois':      mois,
        'annee':     annee,
        'mois_label': MOIS_FR[mois],
        'mois_fr':   MOIS_FR,
    })


# ═══════════════════════════════════════════════════════════════
# SAISIES MEMBRES EN ATTENTE
# ═══════════════════════════════════════════════════════════════


@bureau_required
def saisies_en_attente(request):
    """Liste des saisies soumises par les membres à valider."""
    config = ConfigExercice.get_exercice_courant()
    # Les membres soumettent via ReleveBancaire avec est_valide=False
    from apps.banque.models import ReleveBancaire
    saisies = ReleveBancaire.objects.filter(
        est_valide_membre=True,
        est_valide_bureau=False,
        config_exercice=config
    ).select_related('adherent').order_by('-date_saisie')

    return render(request, 'dashboard/saisie/membres_en_attente.html', {
        'config_exercice': config,
        'saisies':        saisies,
        'nb_en_attente':  saisies.count(),
    })


import json
from django.http import JsonResponse
from django.views.decorators.http import require_POST
from django.views.decorators.csrf import csrf_exempt


@bureau_required
@require_POST
def valider_saisie(request):
    """Valide une saisie membre soumise."""
    try:
        data = json.loads(request.body)
        pk   = data.get('pk')
        from apps.banque.models import ReleveBancaire
        releve = ReleveBancaire.objects.get(pk=pk)
        releve.est_valide_bureau = True
        releve.save()
        # Recalculer MouvementFonds
        saisie = SaisieMonthly.objects.filter(
            adherent=releve.adherent,
            mois=releve.mois,
            annee=releve.annee
        ).first()
        if saisie:
            from apps.saisie.services import calculer_mouvement_fonds
            calculer_mouvement_fonds(saisie)
        return JsonResponse({'ok': True})
    except Exception as e:
        return JsonResponse({'ok': False, 'error': str(e)})


@bureau_required
@require_POST
def rejeter_saisie(request):
    """Rejette et supprime une saisie membre soumise."""
    try:
        data = json.loads(request.body)
        pk   = data.get('pk')
        from apps.banque.models import ReleveBancaire
        ReleveBancaire.objects.filter(pk=pk, est_valide_bureau=False).delete()
        return JsonResponse({'ok': True})
    except Exception as e:
        return JsonResponse({'ok': False, 'error': str(e)})