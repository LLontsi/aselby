"""
PATCH apps/tontines/views.py
Ajouter les vues de saisie tontine mensuelle.
Garder toutes les vues existantes, ajouter après repartition_interets :
"""
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.utils import timezone
from django.db import transaction
from decimal import Decimal
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from .models import NiveauTontine, SessionTontine, ParticipationTontine

MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']
MOIS_CODE = ['','JANV','FEV','MARS','AVRIL','MAI','JUIN',
             'JUIL','AOUT','SEPT','OCT','NOV','DEC']

# ─── vues existantes (inchangées) ─────────────────────────────────────────────
@bureau_required
def tableau_mensuel(request):
    config = ConfigExercice.get_exercice_courant()
    mois = int(request.GET.get('mois', timezone.now().month))
    annee = int(request.GET.get('annee', timezone.now().year))
    niveau_code = request.GET.get('niveau', '')

    sessions = SessionTontine.objects.filter(
        mois=mois, annee=annee, niveau__config_exercice=config
    ).select_related('niveau')
    if not sessions.exists() and not request.GET.get('mois'):
        derniere = SessionTontine.objects.filter(
            niveau__config_exercice=config
        ).order_by('-annee', '-mois').first()
        if derniere:
            mois, annee = derniere.mois, derniere.annee
            sessions = SessionTontine.objects.filter(
                mois=mois, annee=annee, niveau__config_exercice=config
            ).select_related('niveau')

    participations_qs = ParticipationTontine.objects.filter(
        session__mois=mois, session__annee=annee,
        session__niveau__config_exercice=config
    ).select_related('adherent', 'session__niveau').order_by('adherent__numero_ordre')

    niveau_actif = None
    if niveau_code:
        participations_qs = participations_qs.filter(session__niveau__code=niveau_code)
        niveau_actif = NiveauTontine.objects.filter(
            config_exercice=config, code=niveau_code).first()

    # Navigation mois
    prec = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv = (1, annee+1)  if mois == 12 else (mois+1, annee)

    ctx = {
        'config_exercice': config,
        'sessions': sessions,
        'mois': mois, 'annee': annee,
        'mois_label': MOIS_FR[mois],
        'participations': participations_qs,
        'niveau_actif': niveau_actif,
        'niveau_code': niveau_code,
        'niveaux': NiveauTontine.objects.filter(config_exercice=config),
        'prec_mois': prec[0], 'prec_annee': prec[1],
        'suiv_mois': suiv[0], 'suiv_annee': suiv[1],
        'mois_fr': MOIS_FR,
    }
    return render(request, 'dashboard/tontines/tableau.html', ctx)


@bureau_required
def saisie_tontine(request, niveau_code):
    """
    Formulaire de saisie mensuelle pour UN niveau tontine.
    Affiche tous les adhérents du niveau avec leurs champs éditables.
    """
    config = ConfigExercice.get_exercice_courant()
    mois   = int(request.GET.get('mois',  request.POST.get('mois',  timezone.now().month)))
    annee  = int(request.GET.get('annee', request.POST.get('annee', config.annee)))
    niveau = get_object_or_404(NiveauTontine, config_exercice=config, code=niveau_code)

    # Récupérer ou créer la session pour ce mois/niveau
    session, _ = SessionTontine.objects.get_or_create(
        niveau=niveau, mois=mois, annee=annee,
        defaults={'date_seance': timezone.now().date(), 'montant_interet_bureau': Decimal('0')}
    )

    from apps.adherents.models import Adherent
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')

    if request.method == 'POST' and 'save' in request.POST:
        with transaction.atomic():
            for adh in adherents:
                pfx = f"adh_{adh.matricule}_"
                nb_parts = int(request.POST.get(f"{pfx}nombre_parts", 0) or 0)
                if nb_parts == 0:
                    # Pas de participation ce mois
                    ParticipationTontine.objects.filter(
                        session=session, adherent=adh).delete()
                    continue

                mode = request.POST.get(f"{pfx}mode_versement", 'ECHEC')
                num_cheque = request.POST.get(f"{pfx}numero_cheque", '')

                def d(key):
                    val = request.POST.get(f"{pfx}{key}", '0') or '0'
                    try:
                        return Decimal(str(val).replace(' ', ''))
                    except Exception:
                        return Decimal('0')

                # Calcul montant_verse = nb_parts × versement_mensuel_par_part
                montant_attendu = nb_parts * niveau.versement_mensuel_par_part
                montant_verse   = d('montant_verse') or montant_attendu

                # Pénalités
                pen_esp  = config.penalite_especes if mode == 'ESPECES' else Decimal('0')
                pen_ech  = (montant_attendu * config.pourcentage_penalite_echec / 100
                            if mode == 'ECHEC' else Decimal('0'))

                ParticipationTontine.objects.update_or_create(
                    session=session, adherent=adh,
                    defaults=dict(
                        nombre_parts=nb_parts,
                        mode_versement=mode,
                        montant_verse=montant_verse,
                        penalite_especes=pen_esp,
                        penalite_echec=pen_ech,
                        numero_cheque=num_cheque,
                        complement_epargne=d('complement_epargne'),
                        vente_petit_lot=d('vente_petit_lot'),
                        interet_petit_lot=d('interet_petit_lot'),
                        remboursement_petit_lot=d('remboursement_petit_lot'),
                        mode_remboursement_petit_lot=request.POST.get(f"{pfx}mode_remb_petit_lot", ''),
                        montant_lot_principal=d('montant_lot_principal'),
                        interet_lot_principal=d('interet_lot_principal'),
                        a_obtenu_lot_principal=d('montant_lot_principal') > 0,
                    )
                )
        messages.success(request,
            f"Saisie tontine {niveau.code} — {MOIS_FR[mois]} {annee} enregistrée.")
        return redirect(f"?mois={mois}&annee={annee}")

    # GET — préparer les données
    participations = {
        p.adherent_id: p
        for p in ParticipationTontine.objects.filter(
            session=session).select_related('adherent')
    }

    prec = (12, annee-1) if mois == 1 else (mois-1, annee)
    suiv = (1, annee+1)  if mois == 12 else (mois+1, annee)

    ctx = dict(
        config_exercice=config, niveau=niveau, session=session,
        adherents=adherents, participations=participations,
        mois=mois, annee=annee, mois_label=MOIS_FR[mois],
        prec_mois=prec[0], prec_annee=prec[1],
        suiv_mois=suiv[0], suiv_annee=suiv[1],
        mois_fr=MOIS_FR,
    )
    return render(request, 'dashboard/tontines/saisie_tontine.html', ctx)


@bureau_required
def detail_participants(request, niveau_code):
    config = ConfigExercice.get_exercice_courant()
    niveau = get_object_or_404(NiveauTontine, config_exercice=config, code=niveau_code)
    participations = ParticipationTontine.objects.filter(
        session__niveau=niveau
    ).select_related('adherent', 'session').order_by('session__mois', 'adherent__numero_ordre')
    sessions = SessionTontine.objects.filter(niveau=niveau).order_by('mois')
    ctx = {
        'config_exercice': config,
        'niveau': niveau,
        'participations': participations,
        'sessions': sessions,
    }
    return render(request, 'dashboard/tontines/detail_participants.html', ctx)


@bureau_required
def enchere(request):
    config = ConfigExercice.get_exercice_courant()
    return render(request, 'dashboard/tontines/enchere.html', {'config_exercice': config})


@bureau_required
def calendrier(request):
    config = ConfigExercice.get_exercice_courant()
    niveaux = NiveauTontine.objects.filter(config_exercice=config)
    return render(request, 'dashboard/tontines/calendrier.html',
                  {'config_exercice': config, 'niveaux': niveaux})


@bureau_required
def repartition_interets(request):
    config = ConfigExercice.get_exercice_courant()
    mois   = int(request.GET.get('mois', timezone.now().month))
    annee  = int(request.GET.get('annee', timezone.now().year))
    niveaux = NiveauTontine.objects.filter(config_exercice=config)
    ctx = {'config_exercice': config, 'niveaux': niveaux, 'mois': mois, 'annee': annee}
    return render(request, 'dashboard/tontines/interets.html', ctx)


@bureau_required
def telecharger_tontine(request, niveau_code):
    """Génère le fichier Excel TONTINE{taux}.xlsx pour un niveau donné."""
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from django.http import HttpResponse

    config = ConfigExercice.get_exercice_courant()
    niveau = get_object_or_404(NiveauTontine, config_exercice=config, code=niveau_code)
    annee  = config.annee
    wb     = Workbook()
    wb.remove(wb.active)
    BLEU = '1B2B5E'; BLANC = 'FFFFFF'

    COLS = ['MATRICULE','NBREPART','NOM ET PRENOM','MONTANT TONTINE',
            'MONTANT VERSE','PENALITE VERSEMENT ESPECES','MODE VERSEMENT',
            'VENTE PETIT LOT','INTERET PETIT LOT','MODE VERSEMENT VENTE PETIT LOT',
            'NUMERO CHEQUE','MONTANT A REMBOURSER PETIT LOT','REMBOURSEMENT PETIT LOT',
            'MODE VERSEMENT REMBOURSEMENT','MODE VERSEMENTLOT PRINCIPAL',
            'NUMERO CHEQUE (LOT)','MONTANT LOT PRINCIPAL','INTERET LOT PRINCIPAL']

    COLS_R = COLS + ['MOIS OBTENTION','REPARTITION INTERET','INTERET PERCU PAR ADHERENT']

    def _hdr(ws, cols, sheet_name):
        ws.cell(1,1,'ASSOCIATION'); ws.cell(1,2,'ASELBY')
        ws.cell(2,1,'ANNEE'); ws.cell(2,2,annee)
        ws.cell(2,3,'TONTINE'); ws.cell(2,4,f'NBRE PART')
        ws.cell(3,1,'TAUX TONTINE'); ws.cell(3,2,float(niveau.taux_mensuel))
        fill = PatternFill('solid', fgColor=BLEU)
        fn   = Font(bold=True, color=BLANC, size=8, name='Arial')
        for j,c in enumerate(cols,1):
            cell = ws.cell(5,j,c)
            cell.fill=fill; cell.font=fn
            cell.alignment = Alignment(wrap_text=True, horizontal='center')
            ws.column_dimensions[get_column_letter(j)].width = 14
        ws.row_dimensions[5].height = 30

    def _write_row(ws, row_num, adh, part, is_resume=False):
        montant_tontine = part.nombre_parts * niveau.taux_mensuel if part else Decimal('0')
        vals = [
            adh.matricule,
            part.nombre_parts if part else 0,
            adh.nom_prenom,
            float(montant_tontine),
            float(part.montant_verse if part else 0),
            float(part.penalite_especes if part else 0),
            part.mode_versement if part else '',
            float(part.vente_petit_lot if part else 0),
            float(part.interet_petit_lot if part else 0),
            '',  # mode vente petit lot — dans ComplementHistorique
            part.numero_cheque if part else '',
            float(part.remboursement_petit_lot if part else 0),  # montant à rembourser
            float(part.remboursement_petit_lot if part else 0),
            part.mode_remboursement_petit_lot if part else '',
            '',  # mode lot principal — dans ComplementHistorique
            '',  # numéro chèque lot — dans ComplementHistorique
            float(part.montant_lot_principal if part else 0),
            float(part.interet_lot_principal if part else 0),
        ]
        if is_resume:
            vals += ['', '', '']  # mois obtention, repartition interet, interet percu
        for j,v in enumerate(vals,1):
            ws.cell(row_num,j,v)

    from apps.adherents.models import Adherent
    adherents = Adherent.objects.filter(statut='ACTIF').order_by('numero_ordre')

    # Feuilles mensuelles
    for m in range(1, 13):
        label = f'TONT{MOIS_CODE[m]}{str(annee)[-2:]}'
        ws = wb.create_sheet(label)
        _hdr(ws, COLS, label)
        session = SessionTontine.objects.filter(
            niveau=niveau, mois=m, annee=annee).first()
        parts = {}
        if session:
            parts = {p.adherent_id: p for p in
                     ParticipationTontine.objects.filter(session=session)
                     .select_related('adherent')}
        for i, adh in enumerate(adherents, 6):
            _write_row(ws, i, adh, parts.get(adh.matricule))

    # Feuille résumé
    ws_r = wb.create_sheet(f'TONTRESUME{str(annee)[-2:]}')
    _hdr(ws_r, COLS_R, 'RESUME')
    parts_all = {p.adherent_id: p for p in
                 ParticipationTontine.objects.filter(
                     session__niveau=niveau, session__annee=annee)
                 .select_related('adherent')}
    for i, adh in enumerate(adherents, 6):
        _write_row(ws_r, i, adh, parts_all.get(adh.matricule), is_resume=True)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    taux = int(niveau.taux_mensuel)
    resp = HttpResponse(buf.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="ASELBY{annee}TONTINE{taux}.xlsx"'
    return resp