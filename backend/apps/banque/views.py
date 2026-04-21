"""PATCH apps/banque/views.py — ajouter vues TABBHISTOBQUE et TABBORDAIDEDEPENSES"""
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.utils import timezone
from django.db.models import Sum
from django.http import HttpResponse
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from apps.saisie.models import TableauDeBord
from apps.adherents.models import Adherent
from .models import HistoriqueBancaire, Cheque
from decimal import Decimal

MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']
MOIS_CODE = ['','JANV','FEV','MARS','AVRIL','MAI','JUIN',
             'JUIL','AOUT','SEPT','OCT','NOV','DEC']

# ─── Vue existante : historique ────────────────────────────────────────────────
@bureau_required
def historique(request):
    config = ConfigExercice.get_exercice_courant()
    mois  = int(request.GET.get('mois',  timezone.now().month))
    annee = int(request.GET.get('annee', timezone.now().year))

    historiques = HistoriqueBancaire.objects.filter(
        mois=mois, annee=annee, config_exercice=config
    ).select_related('adherent').order_by('adherent__numero_ordre')

    if not historiques.exists() and not request.GET.get('mois'):
        dernier = TableauDeBord.objects.filter(
            config_exercice=config
        ).order_by('-annee', '-mois').first()
        if dernier:
            mois, annee = dernier.mois, dernier.annee

    saisies = TableauDeBord.objects.filter(
        mois=mois, annee=annee, config_exercice=config
    ).select_related('adherent').order_by('adherent__numero_ordre')

    prec = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv = (1, annee+1)  if mois == 12 else (mois+1, annee)

    ctx = {
        'config_exercice': config,
        'historiques':    historiques if historiques.exists() else saisies,
        'saisies':        saisies,
        'depuis_saisies': not historiques.exists(),
        'mois': mois, 'annee': annee,
        'mois_label': MOIS_FR[mois] if 1 <= mois <= 12 else str(mois),
        'prec_mois': prec[0], 'prec_annee': prec[1],
        'suiv_mois': suiv[0], 'suiv_annee': suiv[1],
    }
    return render(request, 'dashboard/banque/historique.html', ctx)


# ─── Nouvelle vue : TABBHISTOBQUE ─────────────────────────────────────────────
@bureau_required
def tabbhistobque(request):
    """
    Vue TABBHISTOBQUE = historique bancaire détaillé par adhérent/mois.
    Saisie + consultation + téléchargement Excel.
    """
    config = ConfigExercice.get_exercice_courant()
    mois   = int(request.GET.get('mois',  timezone.now().month))
    annee  = int(request.GET.get('annee', config.annee))
    prec   = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv   = (1, annee+1)  if mois == 12 else (mois+1, annee)
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')

    if request.method == 'POST' and 'save' in request.POST:
        for adh in adherents:
            pfx = f"adh_{adh.matricule}_"
            def d(k):
                v = request.POST.get(f"{pfx}{k}", '0') or '0'
                try: return Decimal(str(v).replace(' ',''))
                except: return Decimal('0')

            HistoriqueBancaire.objects.update_or_create(
                adherent=adh, mois=mois, annee=annee,
                defaults=dict(
                    config_exercice=config,
                    versement_tontine=d('versement_tontine'),
                    versement_especes=d('versement_especes'),
                    versement_banque=d('versement_banque'),
                    autre_versement=d('autre_versement'),
                    montant_engagement=d('montant_engagement'),
                    agio=d('agio'),
                    en_compte_reel=d('en_compte_reel'),
                    montant_a_justifier_saisi=d('montant_a_justifier_saisi'),
                )
            )
        messages.success(request,
            f"Historique bancaire {MOIS_FR[mois]} {annee} enregistré.")
        return redirect(f"?mois={mois}&annee={annee}")

    # GET — préparer lignes (hist existant ou calculé depuis TableauDeBord)
    hist_map = {
        h.adherent_id: h
        for h in HistoriqueBancaire.objects.filter(
            mois=mois, annee=annee, config_exercice=config)
    }
    tb_map = {
        tb.adherent_id: tb
        for tb in TableauDeBord.objects.filter(
            mois=mois, annee=annee, config_exercice=config)
    }

    rows = []
    for adh in adherents:
        hist = hist_map.get(adh.matricule)
        tb   = tb_map.get(adh.matricule)
        # Préremplir depuis TableauDeBord si pas de HistoriqueBancaire
        if not hist and tb:
            hist = HistoriqueBancaire(
                adherent=adh, mois=mois, annee=annee, config_exercice=config,
                versement_banque=tb.versement_banque,
                versement_especes=tb.versement_especes,
                autre_versement=tb.autre_versement,
                montant_engagement=tb.montant_engagement,
            )
        rows.append({'adherent': adh, 'hist': hist, 'tb': tb})

    # Totaux
    total_tontine  = sum(r['hist'].versement_tontine  if r['hist'] else 0 for r in rows)
    total_banque   = sum(r['hist'].versement_banque   if r['hist'] else 0 for r in rows)
    total_especes  = sum(r['hist'].versement_especes  if r['hist'] else 0 for r in rows)
    total_engagement = sum(r['hist'].montant_engagement if r['hist'] else 0 for r in rows)

    ctx = dict(
        config_exercice=config, rows=rows,
        mois=mois, annee=annee, mois_label=MOIS_FR[mois],
        prec_mois=prec[0], prec_annee=prec[1],
        suiv_mois=suiv[0], suiv_annee=suiv[1],
        total_tontine=total_tontine, total_banque=total_banque,
        total_especes=total_especes, total_engagement=total_engagement,
    )
    return render(request, 'dashboard/banque/tabbhistobque.html', ctx)


# ─── Nouvelle vue : TABBORDAIDEDEPENSES ───────────────────────────────────────
@bureau_required
def tabbordaidedepenses(request):
    """
    Vue calculée TABBORDAIDEDEPENSES = récapitulatif aide/dépenses par adhérent/mois.
    Agrège TableauDeBord + ComplementHistorique + AgioBancaire.
    """
    config = ConfigExercice.get_exercice_courant()
    mois   = int(request.GET.get('mois', timezone.now().month))
    annee  = int(request.GET.get('annee', config.annee))
    prec   = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv   = (1, annee+1)  if mois == 12 else (mois+1, annee)
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')

    # TableauDeBord
    tb_map = {tb.adherent_id: tb for tb in
              TableauDeBord.objects.filter(mois=mois, annee=annee, config_exercice=config)}

    # ComplementHistorique pour num_cheque et montant_cheque_effectif
    try:
        from apps.rapports.models import ComplementHistorique
        ch_map = {ch.adherent_id: ch for ch in
                  ComplementHistorique.objects.filter(
                      tableau_bord__mois=mois, tableau_bord__annee=annee,
                      config_exercice=config).select_related('tableau_bord')}
    except Exception:
        ch_map = {}

    # Agio global du mois
    from apps.parametrage.models import AgioBancaire
    agio_mois = AgioBancaire.objects.filter(
        config_exercice=config, mois=mois, annee=annee).first()
    agio_global = agio_mois.montant_agio if agio_mois else Decimal('0')

    # HistoriqueBancaire pour "extrait de compte"
    hist_map = {h.adherent_id: h for h in
                HistoriqueBancaire.objects.filter(
                    mois=mois, annee=annee, config_exercice=config)}

    rows = []
    total_cheque = total_especes = total_totaux = Decimal('0')
    for adh in adherents:
        tb = tb_map.get(adh.matricule)
        ch = ch_map.get(adh.matricule)
        ht = hist_map.get(adh.matricule)
        montant_cheque = tb.versement_banque if tb else Decimal('0')
        num_cheque     = ch.numero_cheque_effectif if ch else ''
        especes        = tb.versement_especes if tb else Decimal('0')
        totaux         = montant_cheque + especes
        extrait        = ht.en_compte_reel if ht else Decimal('0')

        total_cheque  += montant_cheque
        total_especes += especes
        total_totaux  += totaux

        rows.append({
            'adherent':      adh,
            'totaux':        totaux,
            'montant_cheque': montant_cheque,
            'num_cheque':    num_cheque,
            'especes':       especes,
            'agio':          agio_global,
            'extrait':       extrait,
        })

    ctx = dict(
        config_exercice=config, rows=rows,
        mois=mois, annee=annee, mois_label=MOIS_FR[mois],
        prec_mois=prec[0], prec_annee=prec[1],
        suiv_mois=suiv[0], suiv_annee=suiv[1],
        total_cheque=total_cheque, total_especes=total_especes,
        total_totaux=total_totaux, agio_global=agio_global,
    )
    return render(request, 'dashboard/banque/tabbordaidedepenses.html', ctx)


# ─── Téléchargement TABBHISTOBQUE ─────────────────────────────────────────────
@bureau_required
def telecharger_tabbhistobque(request):
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    config = ConfigExercice.get_exercice_courant()
    annee  = config.annee
    wb     = Workbook()
    wb.remove(wb.active)
    BLEU = '1B2B5E'; BLANC = 'FFFFFF'

    COLS = ['MATRICULE','NOM ET PRENOM','HISTORIQUE TONTINE','HITORIQUE ESPECES',
            'HITORIQUE BANQUE','AUTRE VERSEMENT','EN COMPTE','MONTANT ENGAGEMENT',
            'MONTANT A JUSTIFIER','AGIO']

    def _hdr(ws):
        ws.cell(1,1,'ASSOCIATION'); ws.cell(1,2,'ASELBY')
        ws.cell(2,1,'ANNEE'); ws.cell(2,2,annee)
        ws.cell(3,1,'HISTORIQUE BANCAIRE')
        fill = PatternFill('solid', fgColor=BLEU)
        fn   = Font(bold=True, color=BLANC, size=8, name='Arial')
        for j,c in enumerate(COLS,1):
            cell = ws.cell(5,j,c)
            cell.fill=fill; cell.font=fn
            cell.alignment = Alignment(wrap_text=True, horizontal='center')
            ws.column_dimensions[get_column_letter(j)].width = 16

    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')
    sheets = [(12, annee-1, f'HISTOBQUEDEC{str(annee-1)[-2:]}')]
    for m in range(1,13):
        sheets.append((m, annee, f'HISTOBQUE{MOIS_CODE[m]}{str(annee)[-2:]}'))

    for (m, a, label) in sheets:
        ws = wb.create_sheet(label)
        _hdr(ws)
        hist_map = {h.adherent_id: h for h in
                    HistoriqueBancaire.objects.filter(mois=m, annee=a)}
        tb_map   = {tb.adherent_id: tb for tb in
                    TableauDeBord.objects.filter(mois=m, annee=a)}
        for i, adh in enumerate(adherents, 6):
            h  = hist_map.get(adh.matricule)
            tb = tb_map.get(adh.matricule)
            vt = float(h.versement_tontine if h else 0)
            ve = float((h.versement_especes if h else 0) or (tb.versement_especes if tb else 0))
            vb = float((h.versement_banque if h else 0)  or (tb.versement_banque if tb else 0))
            av = float(h.autre_versement   if h else 0)
            ec = float(h.en_compte         if h else (vt+ve+vb+av))
            me = float((h.montant_engagement if h else 0) or (tb.montant_engagement if tb else 0))
            aj = float(h.montant_a_justifier if h else max(me-(vt+ve+vb+av), 0))
            ag = float(h.agio if h else 0)
            ws.cell(i,1,adh.matricule); ws.cell(i,2,adh.nom_prenom)
            for j,v in enumerate([vt,ve,vb,av,ec,me,aj,ag], 3):
                ws.cell(i,j,v)

    # Feuille résumé
    ws_r = wb.create_sheet(f'HISTOBQUERESUME{str(annee)[-2:]}')
    _hdr(ws_r)
    hist_all = {}
    for h in HistoriqueBancaire.objects.filter(annee=annee).select_related('adherent'):
        hist_all.setdefault(h.adherent_id, []).append(h)
    for i, adh in enumerate(adherents, 6):
        hlist = hist_all.get(adh.matricule, [])
        vt=ve=vb=av=ec=me=aj=ag = 0
        for h in hlist:
            vt+=float(h.versement_tontine); ve+=float(h.versement_especes)
            vb+=float(h.versement_banque);  av+=float(h.autre_versement)
            ec+=float(h.en_compte);         me+=float(h.montant_engagement)
            aj+=float(h.montant_a_justifier); ag+=float(h.agio)
        ws_r.cell(i,1,adh.matricule); ws_r.cell(i,2,adh.nom_prenom)
        for j,v in enumerate([vt,ve,vb,av,ec,me,aj,ag],3):
            ws_r.cell(i,j,v)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    resp = HttpResponse(buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="ASELBY{annee}TABBHISTOBQUE.xlsx"'
    return resp


# ─── Vues existantes (inchangées) ─────────────────────────────────────────────
@bureau_required
def cheques(request):
    config = ConfigExercice.get_exercice_courant()
    mois  = request.GET.get('mois')
    annee = request.GET.get('annee')
    cheques_qs = Cheque.objects.filter(
        config_exercice=config
    ).select_related('adherent').order_by('annee', 'mois', 'adherent__numero_ordre')
    if mois and annee:
        cheques_qs = cheques_qs.filter(mois=int(mois), annee=int(annee))
    mois_disponibles = (Cheque.objects.filter(config_exercice=config)
        .values('mois','annee').order_by('annee','mois').distinct())
    ctx = {
        'config_exercice':  config,
        'cheques':          cheques_qs,
        'mois_filter':      int(mois) if mois else None,
        'annee_filter':     int(annee) if annee else None,
        'mois_disponibles': mois_disponibles,
        'total_cheques':    cheques_qs.aggregate(t=Sum('montant'))['t'] or 0,
    }
    return render(request, 'dashboard/banque/cheques.html', ctx)


@bureau_required
def tresorerie(request):
    config = ConfigExercice.get_exercice_courant()
    saisies_all = TableauDeBord.objects.filter(
        config_exercice=config
    ).values('mois','annee').annotate(
        banque=Sum('versement_banque'),
        especes=Sum('versement_especes'),
        engagement=Sum('montant_engagement'),
    ).order_by('annee','mois')

    synthese_mensuelle = []
    for s in saisies_all:
        en_compte  = (s['banque'] or Decimal('0')) + (s['especes'] or Decimal('0'))
        engagement = s['engagement'] or Decimal('0')
        synthese_mensuelle.append({
            'mois': s['mois'], 'annee': s['annee'],
            'banque': s['banque'] or Decimal('0'),
            'especes': s['especes'] or Decimal('0'),
            'en_compte': en_compte,
            'engagement': engagement,
            'a_justifier': max(engagement - en_compte, Decimal('0')),
        })

    ctx = {
        'config_exercice': config,
        'synthese_mensuelle': synthese_mensuelle,
        'total_banque':  sum(l['banque']    for l in synthese_mensuelle),
        'total_especes': sum(l['especes']   for l in synthese_mensuelle),
        'total_en_compte': sum(l['en_compte'] for l in synthese_mensuelle),
    }
    return render(request, 'dashboard/banque/tresorerie.html', ctx)