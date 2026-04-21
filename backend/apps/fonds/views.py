from django.shortcuts import render, get_object_or_404
from django.utils import timezone
from django.db.models import Sum
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from apps.adherents.models import Adherent
from .models import MouvementFonds, ReserveMensuelle
from .services import calculer_interets_mensuels
from decimal import Decimal

@bureau_required
def etat_mensuel(request):
    config = ConfigExercice.get_exercice_courant()
    mois = int(request.GET.get('mois', timezone.now().month))
    annee = int(request.GET.get('annee', timezone.now().year))
    mouvements = MouvementFonds.objects.filter(mois=mois, annee=annee, config_exercice=config).select_related('adherent').order_by('adherent__numero_ordre')
    # Si aucune donnée pour le mois courant, afficher le dernier mois disponible
    if not mouvements.exists() and not request.GET.get('mois'):
        dernier = MouvementFonds.objects.filter(
            config_exercice=config
        ).order_by('-annee', '-mois').first()
        if dernier:
            mois = dernier.mois
            annee = dernier.annee
            mouvements = MouvementFonds.objects.filter(mois=mois, annee=annee, config_exercice=config).select_related('adherent').order_by('adherent__numero_ordre')
    total_fonds = mouvements.aggregate(t=Sum('fonds_definitif'))['t'] or Decimal('0')
    total_interets = mouvements.aggregate(t=Sum('interet_attribue'))['t'] or Decimal('0')
    nb_eligibles = mouvements.filter(base_calcul_interet__gt=0).count()
    ctx = {'config_exercice': config, 'mouvements': mouvements, 'total_fonds': total_fonds,
           'total_interets': total_interets, 'nb_eligibles': nb_eligibles, 'mois': mois, 'annee': annee}
    return render(request, 'dashboard/fonds/etat.html', ctx)

@bureau_required
def detail_adherent(request, matricule):
    config = ConfigExercice.get_exercice_courant()
    adherent = get_object_or_404(Adherent, matricule=matricule)
    mouvements = MouvementFonds.objects.filter(adherent=adherent, annee=config.annee).order_by('mois')
    ctx = {'config_exercice': config, 'adherent': adherent, 'mouvements': mouvements}
    return render(request, 'dashboard/fonds/detail.html', ctx)

@bureau_required
def repartition_interets(request):
    config = ConfigExercice.get_exercice_courant()
    if request.method == 'POST':
        mois = int(request.POST.get('mois'))
        annee = int(request.POST.get('annee'))
        pool = Decimal(request.POST.get('pool_interets', '0'))
        result = calculer_interets_mensuels(mois, annee, pool, config)
        from django.contrib import messages
        messages.success(request, f"Intérêts répartis : {result['total_distribue']:,.2f} FCFA entre {result['nb_eligibles']} adhérents éligibles.")
    return render(request, 'dashboard/fonds/interets.html', {'config_exercice': config})



# ══════════════════════════════════════════════════════════════
# AJOUTER à la fin de apps/fonds/views.py
# ══════════════════════════════════════════════════════════════

MOIS_CODE = ['','JANV','FEV','MARS','AVRIL','MAI','JUIN',
             'JUIL','AOUT','SEPT','OCT','NOV','DEC']

@bureau_required
def telecharger_listefondscaisse(request):
    """Génère LISTEFONDSCAISSE.xlsx — état du capital de chaque adhérent."""
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from django.http import HttpResponse
    from decimal import Decimal as D

    config    = ConfigExercice.get_exercice_courant()
    annee     = config.annee
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')

    wb   = Workbook(); wb.remove(wb.active)
    BLEU = '1B2B5E'; BLANC = 'FFFFFF'

    def _hdr(ws, cols):
        ws.cell(1, 1, 'MATRICULE'); ws.cell(1, 2, 'ASELBY')
        ws.cell(2, 1, 'ANNEE');     ws.cell(2, 2, annee)
        f  = PatternFill('solid', fgColor=BLEU)
        fn = Font(bold=True, color=BLANC, size=8, name='Arial')
        for j, c in enumerate(cols, 1):
            cl = ws.cell(3, j, c)
            cl.fill = f; cl.font = fn
            cl.alignment = Alignment(wrap_text=True, horizontal='center')
            ws.column_dimensions[get_column_letter(j)].width = 16
        ws.row_dimensions[3].height = 28

    # ── Feuille 1 : FONDSCAISSE (dernier mois disponible) ─────
    ws1 = wb.create_sheet('FONDSCAISSE')
    COLS = ['MATRICULE', 'NUMORDRE', 'NOM ET PRENOM',
            'FONDS DE CAISSE', 'RETRAIT PARTIEL', 'RECONDUCTION', 'TOTAL FONDS']
    _hdr(ws1, COLS)

    # Ligne ASELBY (total)
    total_fonds = MouvementFonds.objects.filter(
        annee=annee, config_exercice=config
    ).aggregate(t=__import__('django.db.models', fromlist=['Sum']).Sum('fonds_definitif'))['t'] or D('0')
    ws1.cell(4, 1, 'AS201648'); ws1.cell(4, 2, 'ASELBY')
    ws1.cell(4, 4, float(total_fonds))

    for i, adh in enumerate(adherents, 5):
        # Dernière saisie
        mvt = MouvementFonds.objects.filter(
            adherent=adh, annee=annee, config_exercice=config
        ).order_by('-mois').first()
        ws1.cell(i, 1, adh.matricule)
        ws1.cell(i, 2, adh.numero_ordre)
        ws1.cell(i, 3, adh.nom_prenom)
        ws1.cell(i, 4, float(mvt.capital_compose_precedent) if mvt else 0)
        ws1.cell(i, 5, float(mvt.retrait_partiel)           if mvt else 0)
        ws1.cell(i, 6, float(mvt.reconduction)              if mvt else 0)
        ws1.cell(i, 7, float(mvt.fonds_definitif)           if mvt else 0)

    # ── Feuille 2 : FONDSTRANSPORT (reconductions) ────────────
    ws2 = wb.create_sheet('FONDSTRANSPORT')
    COLS2 = ['MATRICULE', 'NOM ET PRENOM'] + [MOIS_CODE[m] for m in range(1,13)]
    _hdr(ws2, COLS2)
    for i, adh in enumerate(adherents, 4):
        ws2.cell(i, 1, adh.matricule)
        ws2.cell(i, 2, adh.nom_prenom)
        for m in range(1, 13):
            mvt = MouvementFonds.objects.filter(
                adherent=adh, mois=m, annee=annee, config_exercice=config).first()
            ws2.cell(i, m+2, float(mvt.reconduction) if mvt else 0)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    resp = HttpResponse(
        buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    resp['Content-Disposition'] = f'attachment; filename="ASELBY{annee}LISTEFONDSCAISSE.xlsx"'
    return resp


@bureau_required
def telecharger_basecalculinteret(request):
    """Génère BASECALCULINTERET.xlsx — une feuille par mois avec le calcul des intérêts."""
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from django.http import HttpResponse
    from decimal import Decimal as D

    config    = ConfigExercice.get_exercice_courant()
    annee     = config.annee
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')

    wb   = Workbook(); wb.remove(wb.active)
    BLEU = '1B2B5E'; BLANC = 'FFFFFF'

    COLS = [
        'MATRICULE', 'NOM ET PRENOM',
        'FONDES DE DEPART', 'RECONDUCTION', 'RETRAIT PARTIEL', 'FONDS DEFINITIF',
        'BASE DE CALCUL INTERET FONDS DEFINITIF',
        'REPARTITION PROVISOIRE INTERET FONDS+EPARGNE',
        'CAPITAL COMPOSE',
    ]

    def _hdr(ws, mois):
        ws.cell(1, 1, 'ANNEE'); ws.cell(1, 2, annee if mois >= 1 else annee-1)
        ws.cell(2, 1, 'MOIS');  ws.cell(2, 2, MOIS_CODE[mois] if mois else 'DEC')
        f  = PatternFill('solid', fgColor=BLEU)
        fn = Font(bold=True, color=BLANC, size=8, name='Arial')
        for j, c in enumerate(COLS, 1):
            cl = ws.cell(5, j, c)
            cl.fill = f; cl.font = fn
            cl.alignment = Alignment(wrap_text=True, horizontal='center')
            ws.column_dimensions[get_column_letter(j)].width = 16
        ws.row_dimensions[5].height = 35

    # Feuilles DEC N-1 + JANV→DEC N
    sheets = [(12, annee-1, f'MVTDEC{str(annee-1)[-2:]}')]
    for m in range(1, 13):
        sheets.append((m, annee, f'MVT{MOIS_CODE[m]}{str(annee)[-2:]}'))

    config_prec = None
    try:
        from apps.parametrage.models import ConfigExercice as CE
        config_prec = CE.objects.filter(annee=annee-1).first()
    except Exception:
        pass

    for mois, a, label in sheets:
        ws = wb.create_sheet(label)
        _hdr(ws, mois)
        cfg = config if a == annee else config_prec
        if not cfg:
            continue
        mvts = MouvementFonds.objects.filter(
            mois=mois, annee=a, config_exercice=cfg
        ).select_related('adherent').order_by('adherent__numero_ordre')

        for i, mvt in enumerate(mvts, 6):
            ws.cell(i, 1, mvt.adherent.matricule)
            ws.cell(i, 2, mvt.adherent.nom_prenom)
            ws.cell(i, 3, float(mvt.capital_compose_precedent))
            ws.cell(i, 4, float(mvt.reconduction))
            ws.cell(i, 5, float(mvt.retrait_partiel))
            ws.cell(i, 6, float(mvt.fonds_definitif))
            ws.cell(i, 7, float(mvt.base_calcul_interet))
            ws.cell(i, 8, float(mvt.interet_attribue))
            ws.cell(i, 9, float(mvt.capital_compose))
            # Style gras pour capital composé
            ws.cell(i, 9).font = Font(bold=True, name='Arial', size=9)

        # Ligne totaux
        nb = mvts.count()
        if nb:
            from django.db.models import Sum
            tots = mvts.aggregate(
                fd=Sum('fonds_definitif'), cc=Sum('capital_compose'),
                bi=Sum('base_calcul_interet'), ia=Sum('interet_attribue')
            )
            row_tot = 6 + nb
            ws.cell(row_tot, 1, 'TOTAL')
            ws.cell(row_tot, 6, float(tots['fd'] or 0))
            ws.cell(row_tot, 7, float(tots['bi'] or 0))
            ws.cell(row_tot, 8, float(tots['ia'] or 0))
            ws.cell(row_tot, 9, float(tots['cc'] or 0))
            for j in range(1, 10):
                ws.cell(row_tot, j).font = Font(bold=True, color=BLEU, name='Arial', size=9)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    resp = HttpResponse(
        buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    resp['Content-Disposition'] = f'attachment; filename="ASELBY{annee}BASECALCULINTERET.xlsx"'
    return resp