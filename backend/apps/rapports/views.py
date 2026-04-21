from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.http import HttpResponse
from django.utils import timezone
from django.db.models import Sum
from django.contrib.auth.decorators import login_required
from decimal import Decimal
from apps.parametrage.models import ConfigExercice
from apps.adherents.models import Adherent
from apps.saisie.models import TableauDeBord
from apps.fonds.models import MouvementFonds
from apps.tontines.models import NiveauTontine, ParticipationTontine, SessionTontine
from apps.prets.models import Pret
from apps.dettes.models import ListeRouge
from .models import ComplementMouvement, ComplementHistorique

D = Decimal
MOIS_FR   = ['','Janvier','Février','Mars','Avril','Mai','Juin',
             'Juillet','Août','Septembre','Octobre','Novembre','Décembre']
MOIS_CODE = ['','JANV','FEV','MARS','AVRIL','MAI','JUIN',
             'JUIL','AOUT','SEPT','OCT','NOV','DEC']

# ── Décorateur bureau ──────────────────────────────────────────
def _bureau_required(view_func):
    @login_required
    def wrapper(request, *args, **kwargs):
        if not request.user.est_bureau:
            return redirect('users:mon_espace')
        return view_func(request, *args, **kwargs)
    return wrapper

# ── Dashboard ──────────────────────────────────────────────────
@_bureau_required
def dashboard(request):
    config = ConfigExercice.get_exercice_courant()
    mois   = timezone.now().month
    annee  = timezone.now().year
    nb_actifs = Adherent.objects.filter(statut='ACTIF').count()
    saisies_mois = TableauDeBord.objects.filter(mois=mois, annee=annee, config_exercice=config)
    if not saisies_mois.exists():
        dernier = TableauDeBord.objects.filter(config_exercice=config).order_by('-annee','-mois').first()
        if dernier:
            mois, annee = dernier.mois, dernier.annee
            saisies_mois = TableauDeBord.objects.filter(mois=mois, annee=annee, config_exercice=config)
    total_fonds  = MouvementFonds.objects.filter(annee=annee, mois=mois, config_exercice=config).aggregate(t=Sum('fonds_definitif'))['t'] or D('0')
    prets_qs     = Pret.objects.filter(statut=Pret.EN_COURS)
    total_prets  = prets_qs.aggregate(t=Sum('montant_total_du'))['t'] or D('0')
    demandes_attente = Pret.objects.filter(est_demande_membre=True, est_valide_bureau=False)
    prets_retard     = prets_qs.filter(nb_mois_retard__gt=0)
    nb_saisis = saisies_mois.count()
    alertes = []
    if prets_retard.exists():
        alertes.append({'type':'retard','message':f"{prets_retard.count()} adhérent(s) n\'ont pas remboursé leur prêt ce mois"})
    if demandes_attente.exists():
        alertes.append({'type':'info','message':f"{demandes_attente.count()} demande(s) de prêt en attente"})
    if nb_saisis < nb_actifs:
        alertes.append({'type':'info','message':f"Saisie : {nb_saisis}/{nb_actifs} pour {MOIS_FR[mois]}"})
    else:
        alertes.append({'type':'succes','message':f"Saisie complète pour {MOIS_FR[mois]}"})
    participations_tontines = []
    for niveau in NiveauTontine.objects.filter(config_exercice=config).order_by('taux_mensuel'):
        sess = SessionTontine.objects.filter(niveau=niveau, mois=mois, annee=annee).first()
        nb_parts = ParticipationTontine.objects.filter(session=sess).aggregate(t=Sum('nombre_parts'))['t'] or 0 if sess else 0
        nb_adh   = ParticipationTontine.objects.filter(session=sess).values('adherent').distinct().count() if sess else 0
        participations_tontines.append((niveau.code, nb_parts, nb_adh))
    ctx = {
        'config_exercice': config,
        'nb_adherents': nb_actifs, 'nb_saisis': nb_saisis, 'nb_total': nb_actifs,
        'prets_en_cours': prets_qs.count(),
        'nb_liste_rouge': ListeRouge.objects.filter(est_solde=False).count(),
        'mois_courant': f"{MOIS_FR[mois]} {annee}",
        'entrees_prevues': total_fonds, 'depenses_prevues': D('0'), 'solde_disponible': total_fonds,
        'alertes': alertes,
        'dernieres_saisies': saisies_mois.select_related('adherent').order_by('adherent__numero_ordre')[:10],
        'participations_tontines': participations_tontines,
        'prets_retard': prets_retard.select_related('adherent').order_by('-nb_mois_retard')[:10],
        'demandes_prets': demandes_attente.select_related('adherent')[:5],
        'kpi': {
            'nb_adherents_actifs': nb_actifs, 'total_prets_circulation': total_prets,
            'nb_prets_en_cours': prets_qs.count(),
            'nb_liste_rouge': ListeRouge.objects.filter(est_solde=False).count(),
        },
    }
    return render(request, 'dashboard/dashboard.html', ctx)

# ── Helpers communs ────────────────────────────────────────────
def _prochain_mois_adh(adh, Model, config):
    """Retourne (mois, annee) du prochain mois à saisir pour cet adhérent."""
    # ComplementMouvement : mois/annee via mouvement_fonds FK
    # ComplementHistorique : mois/annee via tableau_bord FK
    if Model == ComplementMouvement:
        last = Model.objects.filter(
            adherent=adh, config_exercice=config
        ).order_by('-mouvement_fonds__annee', '-mouvement_fonds__mois').first()
        if not last:
            return timezone.now().month, config.annee
        m, a = last.mouvement_fonds.mois, last.mouvement_fonds.annee
    else:
        last = Model.objects.filter(
            adherent=adh, config_exercice=config
        ).order_by('-tableau_bord__annee', '-tableau_bord__mois').first()
        if not last:
            return timezone.now().month, config.annee
        m, a = last.tableau_bord.mois, last.tableau_bord.annee
    return (1, a+1) if m == 12 else (m+1, a)

def _save_fields(obj, post, dec_fields, int_fields=None, chr_fields=None):
    for c in dec_fields:
        v = post.get(c,'0').replace(' ','').replace('\xa0','') or '0'
        try: setattr(obj, c, D(str(v)))
        except: pass
    for c in (int_fields or []):
        try: setattr(obj, c, int(post.get(c,0) or 0))
        except: pass
    for c in (chr_fields or []):
        setattr(obj, c, post.get(c,''))

DEC_MVT = ['interet_pret_fonds','remboursement_pret_fonds','penalite_pret_fonds',
    'penalite_fonds','penalite_echec_tontine','penalite_retard_tontine',
    'remboursement_transport','depense_fonds_roulement','depense_frais_exceptionnel',
    'depense_fonds_mutuelle','depense_collation','depense_penalite_banque',
    'pret_definitif','don_volontaire']
CHR_MVT = ['numero_cheque','mode_versement_remboursement','mode_versement_pret',
    'numero_cheque_pret','date_remboursement','statut_pret']

DEC_HISTO = ['epargne_assurance','montant_cheque_effectif','interet_pret',
    'montant_depense','penalite_pret_fonds','remboursement_transport',
    'pret_definitif','nombre_mois_pret']
INT_HISTO = ['nbre_mois_pret']
CHR_HISTO = ['mode_paiement_tontine','num_cheque_versement','mode_paiement_lot',
    'num_cheque_lot','num_cheque_petit_lot','mode_remb_petit_lot',
    'autres_mode_verst','numero_cheque_effectif','date_remboursement','statut_pret']

# ═══════════════════════════════════════════════════════════════
# MOUVEMENTS — LISTE
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def mouvements_liste(request):
    config    = ConfigExercice.get_exercice_courant()
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
    rows = []
    for adh in adherents:
        last = ComplementMouvement.objects.filter(
            adherent=adh, config_exercice=config
        ).order_by('-mouvement_fonds__annee','-mouvement_fonds__mois').first()
        pm, pa = _prochain_mois_adh(adh, ComplementMouvement, config)
        rows.append({'adherent':adh, 'derniere':last, 'prochain_m':pm, 'prochain_a':pa})
    return render(request, 'dashboard/rapports/mouvements_liste.html',
        {'config_exercice':config, 'rows':rows, 'mois_fr':MOIS_FR})

# ═══════════════════════════════════════════════════════════════
# MOUVEMENTS — SAISIE
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def mouvements_saisie(request, matricule):
    config  = ConfigExercice.get_exercice_courant()
    adh     = get_object_or_404(Adherent, matricule=matricule)
    if 'mois' in request.GET or 'mois' in request.POST:
        mois  = int(request.POST.get('mois', request.GET.get('mois')))
        annee = int(request.POST.get('annee', request.GET.get('annee', config.annee)))
    else:
        mois, annee = _prochain_mois_adh(adh, ComplementMouvement, config)

    mvt_obj = MouvementFonds.objects.filter(adherent=adh, mois=mois, annee=annee).first()
    tb_obj  = TableauDeBord.objects.filter(adherent=adh, mois=mois, annee=annee).first()

    if request.method == 'POST' and 'save' in request.POST:
        if not mvt_obj or not tb_obj:
            messages.error(request, f"Saisie normale manquante pour {MOIS_FR[mois]} {annee}.")
            return redirect('rapports:mouvements_liste')
        obj, _ = ComplementMouvement.objects.get_or_create(
            adherent=adh, mouvement_fonds=mvt_obj,
            defaults={'tableau_bord':tb_obj, 'config_exercice':config})
        if not obj.tableau_bord_id:
            obj.tableau_bord = tb_obj
        _save_fields(obj, request.POST, DEC_MVT, chr_fields=CHR_MVT)
        obj.config_exercice = config
        obj.save()
        messages.success(request, f"MVT {adh.nom_prenom} — {MOIS_FR[mois]} {annee} enregistré.")
        return redirect(f"?mois={mois}&annee={annee}")

    obj = ComplementMouvement.objects.filter(
        adherent=adh,
        mouvement_fonds__mois=mois,
        mouvement_fonds__annee=annee
    ).first()
    if not obj and mvt_obj and tb_obj:
        obj = ComplementMouvement(
            adherent=adh, mouvement_fonds=mvt_obj,
            tableau_bord=tb_obj, config_exercice=config)

    tous = list(Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
                .values_list('matricule', flat=True))
    idx  = tous.index(matricule) if matricule in tous else 0
    return render(request, 'dashboard/rapports/mouvements_saisie.html', {
        'config_exercice':config, 'adh':adh, 'obj':obj, 'mvt_obj':mvt_obj,
        'mois':mois, 'annee':annee, 'mois_label':MOIS_FR[mois],
        'adh_prec':tous[idx-1] if idx > 0 else None,
        'adh_suiv':tous[idx+1] if idx < len(tous)-1 else None,
        'mois_fr':MOIS_FR,
    })

# ═══════════════════════════════════════════════════════════════
# MOUVEMENTS — SYNTHÈSE
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def mouvements_synthese(request):
    config = ConfigExercice.get_exercice_courant()
    mois   = max(1, min(12, int(request.GET.get('mois', timezone.now().month))))
    annee  = int(request.GET.get('annee', config.annee))
    cfg    = ConfigExercice.objects.filter(annee=annee).first() or config
    prec   = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv   = (1,  annee+1) if mois == 12 else (mois+1, annee)

    # Tous les adhérents actifs (pas seulement ceux avec ComplementMouvement)
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
    cm_map = {
        cm.adherent_id: cm
        for cm in ComplementMouvement.objects.filter(
            config_exercice=cfg,
            mouvement_fonds__mois=mois,
            mouvement_fonds__annee=annee
        ).select_related('adherent')
    }

    # Construire les lignes — adhérents sans CM auront des valeurs 0/@property via MouvementFonds
    from apps.fonds.models import MouvementFonds
    rows = []
    for adh in adherents:
        cm = cm_map.get(adh.matricule)
        if cm:
            rows.append(cm)
        else:
            # Chercher MouvementFonds pour afficher les données de base
            mvt = MouvementFonds.objects.filter(
                adherent=adh, mois=mois, annee=annee).first()
            if mvt:
                # Créer un objet ComplementMouvement non sauvegardé pour le template
                cm_tmp = ComplementMouvement(
                    adherent=adh, mouvement_fonds=mvt,
                    config_exercice=cfg)
                rows.append(cm_tmp)

    # Totaux sur les vrais CM
    from decimal import Decimal as D
    def tot(attr):
        return sum(getattr(r, attr, D('0')) or D('0') for r in rows if r.pk)

    totaux = {
        'fonds_definitif':        tot('fonds_definitif'),
        'capital_compose':        tot('capital_compose'),
        'reste':                  tot('reste'),
        'epargne':                tot('epargne'),
        'pret_fonds':             tot('pret_fonds'),
        'interet_pret_fonds':     tot('interet_pret_fonds'),
        'remboursement_pret_fonds': tot('remboursement_pret_fonds'),
        'penalite_pret_fonds':    tot('penalite_pret_fonds'),
        'penalite_fonds':         tot('penalite_fonds'),
        'penalite_echec_tontine': tot('penalite_echec_tontine'),
    }

    return render(request, 'dashboard/rapports/mouvements_synthese.html', {
        'config_exercice':config, 'saisies':rows,
        'nb_total': len(adherents), 'nb_saisis': len(cm_map),
        'mois':mois, 'annee':annee, 'mois_label':MOIS_FR[mois],
        'prec_mois':prec[0], 'prec_annee':prec[1], 'prec_label':MOIS_FR[prec[0]],
        'suiv_mois':suiv[0], 'suiv_annee':suiv[1], 'suiv_label':MOIS_FR[suiv[0]],
        'colonnes': ['MATRICULE', 'NOM ET PRENOM', 'FONDES DE DEPART', 'RECONDUCTION', 'RETRAIT PARTIEL', 'FONDS DEFINITIF', 'BASE CALCUL INTERET', 'REPARTITION INTERET', 'CAPITAL COMPOSE', 'SANCTION', 'RESTE', 'EPARGNE', 'FONDS DE ROULEMENT', 'FRAIS EXCEPTIONNEL', 'COLLATION', 'PENALITE VST ESPECES', 'INSCRIPTION', 'MUTUELLE', 'PRET FONDS', 'INTERET PRET FONDS', 'NUMERO CHEQUE', 'REMBOURSEMENT PRET FONDS', 'MODE VST REMBOURSEMENT', 'MODE VST PRET', 'N° CHEQUE PRET', 'PENALITE PRET FONDS', 'PENALITE FONDS', 'PENALITE ECHEC TONTINE', 'PENALITE RETARD TONTINE', 'REMBOURSEMENT TRANSPORT', 'CONTRIBUTION FOYER', 'DEPENSE FONDS ROULEMENT', 'DEPENSE FRAIS EXCEPTIONNEL', 'DEPENSE FONDS MUTUELLE', 'DEPENSE COLLATION', 'DEPENSE PENALITE BANQUE', 'AUTRES DEPENSES'],
        'totaux': totaux,
    })

# ═══════════════════════════════════════════════════════════════
# MOUVEMENTS — RÉSUMÉ
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def mouvements_resume(request):
    from decimal import Decimal as D
    config    = ConfigExercice.get_exercice_courant()
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')

    # Dernière saisie MVT par adhérent
    cm_last = {}
    for cm in ComplementMouvement.objects.filter(
        config_exercice=config
    ).select_related('adherent','mouvement_fonds').order_by(
        'adherent__numero_ordre', '-mouvement_fonds__annee', '-mouvement_fonds__mois'):
        if cm.adherent_id not in cm_last:
            cm_last[cm.adherent_id] = cm

    # Construire les lignes pour TOUS les adhérents
    from apps.fonds.models import MouvementFonds
    saisies = []
    for adh in adherents:
        cm = cm_last.get(adh.matricule)
        if cm:
            saisies.append(cm)
        else:
            # Adhérent sans ComplementMouvement : chercher le dernier MouvementFonds
            mvt = MouvementFonds.objects.filter(
                adherent=adh, config_exercice=config
            ).order_by('-annee', '-mois').first()
            if mvt:
                tb = __import__('apps.saisie.models', fromlist=['TableauDeBord']).TableauDeBord
                tb_obj = tb.objects.filter(
                    adherent=adh,
                    mois=mvt.mois, annee=mvt.annee).first()
                if tb_obj:
                    cm_tmp = ComplementMouvement(
                        adherent=adh, mouvement_fonds=mvt,
                        tableau_bord=tb_obj, config_exercice=config)
                    saisies.append(cm_tmp)

    # Totaux sur les vrais objets sauvegardés
    NUM_FIELDS = [
        'fonds_depart','reconduction','retrait_partiel','fonds_definitif',
        'capital_compose','sanction','reste','epargne','fonds_roulement',
        'frais_exceptionnel','collation','penalite_vst_especes','inscription',
        'mutuelle','pret_fonds','interet_pret_fonds','remboursement_pret_fonds',
        'penalite_pret_fonds','penalite_fonds','penalite_echec_tontine',
        'penalite_retard_tontine','remboursement_transport','contribution_foyer',
        'depense_fonds_roulement','depense_frais_exceptionnel',
        'depense_fonds_mutuelle','depense_collation','depense_penalite_banque',
        'autres_depenses','pret_definitif','don_volontaire',
    ]
    totaux = {}
    for f in NUM_FIELDS:
        totaux[f] = sum(getattr(s, f, D('0')) or D('0') for s in saisies)

    return render(request, 'dashboard/rapports/mouvements_resume.html', {
        'config_exercice':config, 'saisies':saisies, 'annee':config.annee,
        'totaux': totaux,
    })

# ═══════════════════════════════════════════════════════════════
# HISTORIQUE — LISTE
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def historique_liste(request):
    config    = ConfigExercice.get_exercice_courant()
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
    rows = []
    for adh in adherents:
        last = ComplementHistorique.objects.filter(
            adherent=adh, config_exercice=config
        ).order_by('-tableau_bord__annee','-tableau_bord__mois').first()
        pm, pa = _prochain_mois_adh(adh, ComplementHistorique, config)
        rows.append({'adherent':adh, 'derniere':last, 'prochain_m':pm, 'prochain_a':pa})
    return render(request, 'dashboard/rapports/historique_liste.html',
        {'config_exercice':config, 'rows':rows, 'mois_fr':MOIS_FR})

# ═══════════════════════════════════════════════════════════════
# HISTORIQUE — SAISIE
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def historique_saisie(request, matricule):
    config  = ConfigExercice.get_exercice_courant()
    adh     = get_object_or_404(Adherent, matricule=matricule)
    if 'mois' in request.GET or 'mois' in request.POST:
        mois  = int(request.POST.get('mois', request.GET.get('mois')))
        annee = int(request.POST.get('annee', request.GET.get('annee', config.annee)))
    else:
        mois, annee = _prochain_mois_adh(adh, ComplementHistorique, config)

    tb_obj = TableauDeBord.objects.filter(adherent=adh, mois=mois, annee=annee).first()

    if request.method == 'POST' and 'save' in request.POST:
        if not tb_obj:
            messages.error(request, f"Saisie normale manquante pour {MOIS_FR[mois]} {annee}.")
            return redirect('rapports:historique_liste')
        obj, _ = ComplementHistorique.objects.get_or_create(
            adherent=adh, tableau_bord=tb_obj,
            defaults={'config_exercice':config})
        _save_fields(obj, request.POST, DEC_HISTO, INT_HISTO, CHR_HISTO)
        obj.config_exercice = config
        obj.save()
        messages.success(request, f"HISTO {adh.nom_prenom} — {MOIS_FR[mois]} {annee} enregistré.")
        return redirect(f"?mois={mois}&annee={annee}")

    obj = ComplementHistorique.objects.filter(
        adherent=adh,
        tableau_bord__mois=mois,
        tableau_bord__annee=annee
    ).first()
    if not obj and tb_obj:
        obj = ComplementHistorique(
            adherent=adh, tableau_bord=tb_obj, config_exercice=config)

    tous = list(Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
                .values_list('matricule', flat=True))
    idx  = tous.index(matricule) if matricule in tous else 0
    return render(request, 'dashboard/rapports/historique_saisie.html', {
        'config_exercice':config, 'adh':adh, 'obj':obj,
        'mois':mois, 'annee':annee, 'mois_label':MOIS_FR[mois],
        'adh_prec':tous[idx-1] if idx > 0 else None,
        'adh_suiv':tous[idx+1] if idx < len(tous)-1 else None,
        'mois_fr':MOIS_FR,
    })

# ═══════════════════════════════════════════════════════════════
# HISTORIQUE — SYNTHÈSE
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def historique_synthese(request):
    config = ConfigExercice.get_exercice_courant()
    mois   = max(1, min(12, int(request.GET.get('mois', timezone.now().month))))
    annee  = int(request.GET.get('annee', config.annee))
    cfg    = ConfigExercice.objects.filter(annee=annee).first() or config
    prec   = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv   = (1,  annee+1) if mois == 12 else (mois+1, annee)
    saisies = ComplementHistorique.objects.filter(
        config_exercice=cfg,
        tableau_bord__mois=mois,
        tableau_bord__annee=annee
    ).select_related('adherent').order_by('adherent__numero_ordre')
    return render(request, 'dashboard/rapports/historique_synthese.html', {
        'config_exercice':config, 'saisies':saisies,
        'mois':mois, 'annee':annee, 'mois_label':MOIS_FR[mois],
        'prec_mois':prec[0], 'prec_annee':prec[1], 'prec_label':MOIS_FR[prec[0]],
        'suiv_mois':suiv[0], 'suiv_annee':suiv[1], 'suiv_label':MOIS_FR[suiv[0]],
    })

# ═══════════════════════════════════════════════════════════════
# HISTORIQUE — RÉSUMÉ
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def historique_resume(request):
    config  = ConfigExercice.get_exercice_courant()
    vus, saisies = set(), []
    for s in ComplementHistorique.objects.filter(
        config_exercice=config
    ).select_related('adherent').order_by(
        'adherent__numero_ordre', '-tableau_bord__mois'):
        if s.adherent_id not in vus:
            vus.add(s.adherent_id)
            saisies.append(s)
    return render(request, 'dashboard/rapports/historique_resume.html', {
        'config_exercice':config, 'saisies':saisies, 'annee':config.annee})

# ═══════════════════════════════════════════════════════════════
# TÉLÉCHARGEMENTS EXCEL
# ═══════════════════════════════════════════════════════════════
@_bureau_required
def telecharger_mouvements(request):
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    config = ConfigExercice.get_exercice_courant()
    annee  = config.annee
    wb = Workbook(); wb.remove(wb.active)
    BLEU='1B2B5E'; BLANC='FFFFFF'
    COLS=['MATRICULE','NOM ET PRENOM','FONDES DE DEPART','RECONDUCTION',
        'RETRAIT PARTIEL','FONDS DEFINITIF','BASE CALCUL INTERET','REPARTITION INTERET',
        'CAPITAL COMPOSE','SANCTION','RESTE','EPARGNE','FONDS DE ROULEMENT',
        'FRAIS EXCEPTIONNEL','COLLATION','PENALITE VST ESPECES','INSCRIPTION','MUTUELLE',
        'PRET FONDS','INTERET PRET FONDS','NUMERO CHEQUE','REMBOURSEMENT PRET FONDS',
        'MODE VERSEMENT REMBOURSEMENT','MODE VERSEMENT PRET','NUMERO CHEQUE PRET',
        'PENALITE PRET FONDS','PENALITE FONDS','PENALITE ECHEC TONTINE',
        'PENALITE RETARD TONTINE','REMBOURSEMENT TRANSPORT','CONTRIBUTION FOYER',
        'DEPENSE FONDS ROULEMENT','DEPENSE FRAIS EXCEPTIONNEL','DEPENSE FONDS MUTUELLE',
        'DEPENSE COLLATION','DEPENSE PENALITE BANQUE','AUTRES DEPENSES']
    COLS_R = COLS[:21] + ['PRET DEFINITIF','DATE REMBOURSEMENT','STATUT PRET'] + COLS[21:] + ['DON VOLONTAIRE']
    def _hdr(ws, cols):
        ws.cell(1,1,'ASSOCIATION'); ws.cell(1,2,'ASELBY')
        ws.cell(2,1,'ANNEE'); ws.cell(2,2,annee)
        f=PatternFill('solid',fgColor=BLEU); fn=Font(bold=True,color=BLANC,size=8,name='Arial')
        for j,c in enumerate(cols,1):
            cl=ws.cell(5,j,c); cl.fill=f; cl.font=fn
            cl.alignment=Alignment(wrap_text=True,horizontal='center')
            ws.column_dimensions[get_column_letter(j)].width=14
        ws.row_dimensions[5].height=30
    def _row(ws, r, s, is_r=False):
        v=[s.adherent.matricule, s.adherent.nom_prenom,
           float(s.fonds_depart), float(s.reconduction), float(s.retrait_partiel),
           float(s.fonds_definitif), float(s.base_calcul_interet), float(s.repartition_interet),
           float(s.capital_compose), float(s.sanction), float(s.reste), float(s.epargne),
           float(s.fonds_roulement), float(s.frais_exceptionnel), float(s.collation),
           float(s.penalite_vst_especes), float(s.inscription), float(s.mutuelle),
           float(s.pret_fonds), float(s.interet_pret_fonds), s.numero_cheque]
        if is_r:
            v += [float(s.pret_definitif), s.date_remboursement, s.statut_pret]
        v += [float(s.remboursement_pret_fonds), s.mode_versement_remboursement,
              s.mode_versement_pret, s.numero_cheque_pret, float(s.penalite_pret_fonds),
              float(s.penalite_fonds), float(s.penalite_echec_tontine),
              float(s.penalite_retard_tontine), float(s.remboursement_transport),
              float(s.contribution_foyer), float(s.depense_fonds_roulement),
              float(s.depense_frais_exceptionnel), float(s.depense_fonds_mutuelle),
              float(s.depense_collation), float(s.depense_penalite_banque),
              float(s.autres_depenses)]
        if is_r:
            v.append(float(s.don_volontaire))
        for j,val in enumerate(v,1): ws.cell(r,j,val)
    config_prec = ConfigExercice.objects.filter(annee=annee-1).first()
    sheets = [(12, annee-1, f'MVTDEC{str(annee-1)[-2:]}')]
    for m in range(1,13): sheets.append((m, annee, f'MVT{MOIS_CODE[m]}{str(annee)[-2:]}'))
    for m, a, label in sheets:
        ws = wb.create_sheet(label); _hdr(ws, COLS)
        cfg = config if a == annee else config_prec
        if cfg:
            for i,s in enumerate(ComplementMouvement.objects.filter(
                config_exercice=cfg, mouvement_fonds__mois=m, mouvement_fonds__annee=a
            ).select_related('adherent').order_by('adherent__numero_ordre'), 6):
                _row(ws, i, s)
    ws_r = wb.create_sheet(f'MVTRESUME{str(annee)[-2:]}'); _hdr(ws_r, COLS_R)
    vus = set()
    row_num = 6
    for s in ComplementMouvement.objects.filter(config_exercice=config).select_related('adherent').order_by('adherent__numero_ordre','-mouvement_fonds__mois'):
        if s.adherent_id not in vus:
            vus.add(s.adherent_id); _row(ws_r, row_num, s, is_r=True); row_num += 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    resp = HttpResponse(buf.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="ASELBY{annee}AUTREMOUVEMENT.xlsx"'
    return resp


@_bureau_required
def telecharger_historique(request):
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    config = ConfigExercice.get_exercice_courant()
    annee  = config.annee
    wb = Workbook(); wb.remove(wb.active)
    BLEU='1B2B5E'; BLANC='FFFFFF'
    COLS=['MATRICULE','NOM ET PRENOM','BONUS MALUS','VERSEMENT BANQUE','VERSEMENT ESPECES',
        'AUTRE VERSEMENTS','COMPLEMENT EPARGNE','PENALITE VERSEMENT ESPECES',
        'MODE PAIEMENT TONTINE','NBRE LOT T35','TONTINE 35 000','NBRE LOT T75','TONTINE 75 000',
        'NBRE LOT T100','TONTINE 100 000','REMB PETIT LOT T35','REMB PETIT LOT T75','REMB PETIT LOT T100',
        'REMBOURSEMENT PRET FONDS','MODE PAIEMENT REMB PRET','MONTANT ENGAGEMENT','SANCTION','INSCRIPTION',
        'EPARGNE ASSURANCE','TONTINE MOIS','PRET FONDS','MODE PAIEMENT PRET','NUM CHEQUE VERSEMENT',
        'ACHAT LOT T35','ACHAT LOT T75','ACHAT LOT T100','MODE PAIEMENT LOT','NUM CHEQUE LOT',
        'PENALITE VERSEMENT ESPECES','VENTE PETIT LOT T35','VENTE PETIT LOT T75','VENTE PETIT LOT T100',
        'INTERET PETIT LOT T35','INTERET PETIT LOT T75','INTERET PETIT LOT T100',
        'NUM CHEQUE PETIT LOT','MODE REMB PETIT LOT','PENALITE RETARD TONTINE','PENALITE ECHEC TONTINE',
        'RESTE','MUTUELLE','REMBOURSEMENT TRANSPORT','RETRAIT PARTIEL FONDS',
        'MONTANT T25','MONTANT T75','MONTANT T100','MONTANT CHEQUE','MONTANT ESPECES',
        'CONTRIBUTION FOYER','AUTRES MODE VERST','MONTANT CHEQUE EFFECTIF','NUMERO CHEQUE',
        'INTERET PRET','NBRE MOIS','DEPENSE FONDS ROULEMENT','DEPENSE FRAIS EXCEPTIONNEL',
        'DEPENSE FONDS MUTUEL','DEPENSE COLLATION','PENALITE PRET FONDS','DEPENSE PENALITE BANQUE',
        'AUTRES DEPENSES','LIBELLE','COMPTE','MONTANT','DON VOLONTAIRE','PRET DEFINITIF']
    COLS_R = COLS + ['MONTANT PRET DEFINITIF','NOMBRE DE MOIS','DATE REMBOURSEMENT','STATUT PRET']
    def _hdr(ws, cols):
        ws.cell(1,1,'ASSOCIATION'); ws.cell(1,2,'ASELBY')
        ws.cell(2,1,'ANNEE'); ws.cell(2,2,annee)
        f=PatternFill('solid',fgColor=BLEU); fn=Font(bold=True,color=BLANC,size=8,name='Arial')
        for j,c in enumerate(cols,1):
            cl=ws.cell(5,j,c); cl.fill=f; cl.font=fn
            cl.alignment=Alignment(wrap_text=True,horizontal='center')
            ws.column_dimensions[get_column_letter(j)].width=14
        ws.row_dimensions[5].height=30
    def _row(ws, r, s, is_r=False):
        v=[s.adherent.matricule, s.adherent.nom_prenom,
           float(s.bonus_malus), float(s.versement_banque), float(s.versement_especes),
           float(s.autre_versement), float(s.complement_epargne), float(s.penalite_vst_especes),
           s.mode_paiement_tontine, s.nbre_lot_t35, float(s.tontine_35),
           s.nbre_lot_t75, float(s.tontine_75), s.nbre_lot_t100, float(s.tontine_100),
           float(s.remb_petit_lot_t35), float(s.remb_petit_lot_t75), float(s.remb_petit_lot_t100),
           float(s.remboursement_pret), s.mode_paiement_remb_pret,
           float(s.montant_engagement), float(s.sanction), float(s.inscription),
           float(s.epargne_assurance), float(s.tontine_mois), float(s.pret_fonds),
           getattr(s,"mode_paiement_pret","") or getattr(s,"mode_paiement_remb_pret",""), s.num_cheque_versement,
           float(s.achat_lot_t35), float(s.achat_lot_t75), float(s.achat_lot_t100),
           s.mode_paiement_lot, s.num_cheque_lot, float(s.penalite_vst_especes),
           float(s.vente_petit_lot_t35), float(s.vente_petit_lot_t75), float(s.vente_petit_lot_t100),
           float(s.interet_petit_lot_t35), float(s.interet_petit_lot_t75), float(s.interet_petit_lot_t100),
           s.num_cheque_petit_lot, s.mode_remb_petit_lot,
           float(s.penalite_retard_tontine), float(s.penalite_echec_tontine),
           float(s.reste), float(s.mutuelle), float(s.remboursement_transport),
           float(s.retrait_partiel), float(s.montant_t25), float(s.montant_t75), float(s.montant_t100),
           float(s.montant_cheque), float(s.montant_especes), float(s.contribution_foyer),
           s.autres_mode_verst, float(s.montant_cheque_effectif), s.numero_cheque,
           float(s.interet_pret), s.nbre_mois_pret,
           float(s.depense_fonds_roulement), float(s.depense_frais_excep),
           float(s.depense_fonds_mutuel), float(s.depense_collation),
           float(s.penalite_pret_fonds), float(s.depense_penalite_banque),
           float(s.autres_depenses), s.libelle_depense, s.compte_depense,
           float(s.montant_depense), float(s.don_foyer_volontaire), float(s.pret_definitif)]
        if is_r:
            v += [float(s.pret_definitif), s.nombre_mois_pret, s.date_remboursement, s.statut_pret]
        for j,val in enumerate(v,1): ws.cell(r,j,val)
    config_prec = ConfigExercice.objects.filter(annee=annee-1).first()
    sheets = [(12, annee-1, f'HISTODEC{str(annee-1)[-2:]}')]
    for m in range(1,13): sheets.append((m, annee, f'HISTO{MOIS_CODE[m]}{str(annee)[-2:]}'))
    for m, a, label in sheets:
        ws = wb.create_sheet(label); _hdr(ws, COLS)
        cfg = config if a == annee else config_prec
        if cfg:
            for i,s in enumerate(ComplementHistorique.objects.filter(
                config_exercice=cfg, tableau_bord__mois=m, tableau_bord__annee=a
            ).select_related('adherent').order_by('adherent__numero_ordre'), 6):
                _row(ws, i, s)
    ws_r = wb.create_sheet(f'HISTORESUME{str(annee)[-2:]}'); _hdr(ws_r, COLS_R)
    vus = set()
    row_num = 6
    for s in ComplementHistorique.objects.filter(config_exercice=config).select_related('adherent').order_by('adherent__numero_ordre','-tableau_bord__mois'):
        if s.adherent_id not in vus:
            vus.add(s.adherent_id); _row(ws_r, row_num, s, is_r=True); row_num += 1
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    resp = HttpResponse(buf.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="ASELBY{annee}TABBORD.xlsx"'
    return resp