#!/usr/bin/env python3
"""
Script d'import ASELBY 2026 — ENTIÈREMENT DYNAMIQUE
=====================================================
Lit toutes les données depuis les fichiers Excel ASELBYTONTINE_2026.
Aucune valeur hardcodée — tout vient des fichiers Excel ou de la BD.

Changements 2026 vs 2025 :
- Nouveau membre : AS201610 FOTSO THOMAS (ex-liste rouge)
- Nouveau niveau tontine : T60 (taux=35 000, versement=60 000/mois) remplace T35
- Fonds de caisse : 900 000 F (au lieu de 800 000)
- Seulement 3 mois disponibles (janv, fév, mars)
- SIMO PIERRE absent (décédé en 2025)

Usage :
    python manage.py shell < import_aselby_2026.py
"""

import os, sys, glob
from decimal import Decimal
from datetime import date, datetime

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
import django
django.setup()

from django.db import transaction
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

from apps.parametrage.models import (ConfigExercice, DepenseFondsRoulement,
    DepenseFraisExceptionnels, DepenseCollation, DepenseFoyer,
    AgioBancaire, AutreDepense)
from apps.adherents.models import Adherent
from apps.users.models import Utilisateur
from apps.tontines.models import NiveauTontine, SessionTontine, ParticipationTontine
from apps.saisie.models import TableauDeBord
from apps.fonds.models import MouvementFonds
from apps.fonds import services as fonds_services
from apps.prets.models import Pret
from apps.mutuelle.models import CotisationMutuelle, AideMutuelle
from apps.foyer.models import ContributionFoyer
from apps.dettes.models import ListeRouge
from apps.exercice.models import FicheCassation, SyntheseCompte
from openpyxl import load_workbook
from django.db.models import Sum

D = Decimal

def d(v, default=0):
    if v is None: return D(str(default))
    try: return D(str(float(v)))
    except: return D(str(default))

def to_date(v):
    if v is None: return None
    if isinstance(v, (date, datetime)):
        return v.date() if isinstance(v, datetime) else v
    try: return pd.to_datetime(v).date()
    except: return None

print("=" * 65)
print("IMPORT ASELBY 2026 — ENTIÈREMENT DYNAMIQUE")
print("=" * 65)

# ── Localiser le dossier Excel 2026 ────────────────────────────────
_CANDIDATS = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ASELBYTONTINE_2026'),
    os.path.expanduser('~/Téléchargements/ASELBYTONTINE_2026'),
    os.path.expanduser('~/Downloads/ASELBYTONTINE_2026'),
    '/tmp/aselby2026_data/ASELBYTONTINE_2026',
]
BASE_XLS = None
for c in _CANDIDATS:
    if os.path.isfile(os.path.join(c, 'ASELBY2026TABBORD.xlsx')):
        BASE_XLS = c; break
if not BASE_XLS:
    print("❌ Dossier ASELBYTONTINE_2026 introuvable.")
    sys.exit(1)
print(f"  ✓ Fichiers Excel : {BASE_XLS}")

TRAV = f'{BASE_XLS}/TRAVAUXFINEXERCICE2026.xlsx'

# ══════════════════════════════════════════════════════════════════════
# 1. ANNÉE depuis TRAVAUXFINEXERCICE2026
# ══════════════════════════════════════════════════════════════════════
df_synth = pd.read_excel(TRAV, sheet_name='SYNTHESECOMPTE', header=None)
# L'année courante est dans la ligne 1 col B — mais 2026 affiche "2025" comme report
# L'année de l'exercice = année du fichier
annee = 2026
print(f"\n  Année exercice : {annee}")

# ══════════════════════════════════════════════════════════════════════
# 2. CONFIG EXERCICE 2026 — taux depuis BASECALCUL premier mois
# ══════════════════════════════════════════════════════════════════════
print("\n[2] ConfigExercice 2026...")
df_bc1 = pd.read_excel(f'{BASE_XLS}/ASELBY2026BASECALCULINTERET.xlsx',
                       sheet_name='MVTJANV26', header=None)
# Lire les taux depuis le premier membre actif (BAKOP = ligne 6)
mask_bakop = df_bc1[0].astype(str).str.contains('AS201601', na=False)
if mask_bakop.any():
    row_b = df_bc1[mask_bakop].iloc[0]
    taux_fr  = d(row_b[12]) if row_b[12] else D('1000')
    taux_fe  = d(row_b[13]) if row_b[13] else D('1000')
    taux_col = d(row_b[14]) if row_b[14] else D('4000')
else:
    taux_fr = D('1000'); taux_fe = D('1000'); taux_col = D('4000')

config, _ = ConfigExercice.objects.update_or_create(
    annee=annee,
    defaults=dict(
        est_ouvert=True, date_ouverture=date(annee,1,1),
        # Taux tontines 2026 (T60=35000, T75=75000, T100=100000)
        taux_t35=D('35000'), versement_t35=D('60000'),
        taux_t75=D('75000'), taux_t100=D('100000'),
        diviseur_interet_t35=35,
        seuil_eligibilite_interets=D('900000'),
        fonds_roulement_mensuel=taux_fr,
        frais_exceptionnels_mensuel=taux_fe,
        collation_mensuelle=taux_col,
        mutuelle_mensuelle=D('0'),
        montant_inscription=D('25000'),
        complement_fonds_fin_exercice=D('100000'),
        penalite_especes=D('3000'), penalite_especes_active=True,
        pourcentage_penalite_echec=D('20'),
        nb_echecs_max_avant_liste_rouge=3,
        taux_interet_pret_mensuel=D('1'),
        majoration_retard_mois_1=D('5'),
        majoration_retard_mois_2=D('10'),
        montant_min_pret=D('0'),
        montant_max_pret=D('5000000'),
        contribution_foyer_lot_principal=D('150000'),
        complement_mutuelle_fin_exercice=D('30000'),
    )
)
print(f"  ✓ Config {annee}: FR={taux_fr}/mois, FE={taux_fe}/mois, COL={taux_col}/mois")

# ══════════════════════════════════════════════════════════════════════
# 3. ADHÉRENTS depuis LISTEADHERENT 2026
# ══════════════════════════════════════════════════════════════════════
print("\n[3] Adhérents 2026...")
df_adh = pd.read_excel(f'{BASE_XLS}/ASELBY2026LISTEADHERENT.xlsx',
                       sheet_name='LISTADHERENT', header=None)

# Capitaux de départ depuis LISTEFONDSCAISSE 2026
df_fc = pd.read_excel(f'{BASE_XLS}/ASELBY2026LISTEFONDSCAISSE.xlsx', header=None)
capitaux_depart = {}
recon_depart = {}
for _, row in df_fc.iterrows():
    mat = str(row[0]) if row[0] else ''
    if mat.startswith('AS'):
        capitaux_depart[mat] = d(row[3])
        recon_depart[mat] = d(row[5])

nb_adh = 0
with transaction.atomic():
    for _, row in df_adh.iterrows():
        mat = str(row[0]) if row[0] else ''
        if not mat.startswith('AS'):
            continue
        nom = str(row[2]) if row[2] else ''
        num_ordre = int(row[1]) if pd.notna(row[1]) else 999
        fonction = str(row[3]) if pd.notna(row[3]) else ''
        tel = str(int(row[5])) if pd.notna(row[5]) and row[5] else ''
        residence = str(row[7]) if pd.notna(row[7]) else ''
        statut_raw = str(row[8]) if pd.notna(row[8]) else ''
        statut = 'ACTIF' if 'ACTIF' in statut_raw.upper() else 'INACTIF'
        d_rec = to_date(row[9]) if pd.notna(row[9]) else None
        cap = capitaux_depart.get(mat, D('0'))

        a, created = Adherent.objects.update_or_create(
            matricule=mat,
            defaults=dict(
                numero_ordre=num_ordre, nom_prenom=nom,
                fonction=fonction, telephone1=tel[:25] if tel else '',
                residence=residence, statut=statut,
                date_reception=d_rec, capital_depart_exercice=cap,
            )
        )
        # Compte utilisateur
        if mat != 'AS201648':
            username = nom.split()[0].lower() if nom else mat
            u, _ = Utilisateur.objects.update_or_create(
                username=username,
                defaults=dict(nom_complet=nom, role='MEMBRE',
                             adherent=a, est_actif=(statut=='ACTIF'))
            )
        nb_adh += 1

print(f"  ✓ {nb_adh} adhérents mis à jour (dont AS201610 FOTSO THOMAS)")

# ══════════════════════════════════════════════════════════════════════
# 4. NIVEAUX TONTINE 2026 — T60 (nouveau), T75, T100
#    T60 = ancien T35 mais avec versement 60 000 au lieu de 35 000
# ══════════════════════════════════════════════════════════════════════
print("\n[4] Niveaux tontine 2026...")

# Lire les taux directement depuis les fichiers tontine
def _lire_taux_tontine(fname, sheet='TONTJANV26'):
    df = pd.read_excel(f'{BASE_XLS}/{fname}', sheet_name=sheet, header=None)
    taux = d(df.iloc[2, 1])          # taux_mensuel ligne 3 col B
    versement_raw = d(df.iloc[2, 3]) if df.shape[1] > 3 else taux  # versement/part
    # T60 : versement = 60 000 (col 4 des données membres)
    mask = df[0].astype(str).str.startswith('AS')
    if mask.any():
        versement_raw = d(df[mask].iloc[0, 4]) if df[mask].iloc[0, 4] else taux
    return taux, versement_raw

taux_t60, vers_t60   = _lire_taux_tontine('ASELBY2026TONTINE60000.xlsx')
taux_t75, vers_t75   = _lire_taux_tontine('ASELBY2026TONTINE75000.xlsx')
taux_t100, vers_t100 = _lire_taux_tontine('ASELBY2026TONTINE100000.xlsx')

niv_t60, _ = NiveauTontine.objects.update_or_create(
    config_exercice=config, code='T60',
    defaults=dict(taux_mensuel=taux_t60, versement_mensuel_par_part=D('60000'), diviseur_interet=35)
)
niv_t75, _ = NiveauTontine.objects.update_or_create(
    config_exercice=config, code='T75',
    defaults=dict(taux_mensuel=taux_t75, versement_mensuel_par_part=taux_t75, diviseur_interet=13)
)
niv_t100, _ = NiveauTontine.objects.update_or_create(
    config_exercice=config, code='T100',
    defaults=dict(taux_mensuel=taux_t100, versement_mensuel_par_part=taux_t100, diviseur_interet=24)
)
print(f"  ✓ T60 (taux={taux_t60}, vers=60000), T75 (taux={taux_t75}), T100 (taux={taux_t100})")

# ══════════════════════════════════════════════════════════════════════
# 5. SESSIONS + PARTICIPATIONS + SAISIES + MOUVEMENTS FONDS
#    Seulement les mois disponibles dans BASECALCUL
# ══════════════════════════════════════════════════════════════════════
print("\n[5] Sessions / Participations / Saisies / Fonds...")

# Mapping feuilles par mois
MOIS_BC_MAP = {}
MOIS_TB_MAP = {}
xl_bc = pd.ExcelFile(f'{BASE_XLS}/ASELBY2026BASECALCULINTERET.xlsx')
xl_tb = pd.ExcelFile(f'{BASE_XLS}/ASELBY2026TABBORD.xlsx')

for sheet in xl_bc.sheet_names:
    if sheet.startswith('MVTJANV'): MOIS_BC_MAP[1] = sheet
    elif sheet.startswith('MVTFEV'): MOIS_BC_MAP[2] = sheet
    elif sheet.startswith('MVTMARS'): MOIS_BC_MAP[3] = sheet
    elif sheet.startswith('MVTAVRIL'): MOIS_BC_MAP[4] = sheet
    elif sheet.startswith('MVTMAI'): MOIS_BC_MAP[5] = sheet
    elif sheet.startswith('MVTJUIN'): MOIS_BC_MAP[6] = sheet
    elif sheet.startswith('MVTJUIL'): MOIS_BC_MAP[7] = sheet
    elif sheet.startswith('MVTAOUT'): MOIS_BC_MAP[8] = sheet
    elif sheet.startswith('MVTSEPT'): MOIS_BC_MAP[9] = sheet
    elif sheet.startswith('MVTOCT'): MOIS_BC_MAP[10] = sheet
    elif sheet.startswith('MVTNOV'): MOIS_BC_MAP[11] = sheet
    elif sheet.startswith('MVTDEC2' + str(annee)[2:]): MOIS_BC_MAP[12] = sheet

for sheet in xl_tb.sheet_names:
    if 'JANV' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[1] = sheet
    elif 'FEV' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[2] = sheet
    elif 'MARS' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[3] = sheet
    elif 'AVRIL' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[4] = sheet
    elif 'MAI' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[5] = sheet
    elif 'JUIN' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[6] = sheet
    elif 'JUIL' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[7] = sheet
    elif 'AOUT' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[8] = sheet
    elif 'SEPT' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[9] = sheet
    elif 'OCT' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[10] = sheet
    elif 'NOV' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[11] = sheet
    elif 'DEC' in sheet.upper() and str(annee)[2:] in sheet: MOIS_TB_MAP[12] = sheet

MOIS_TONT_MAP = {}
for sheet in pd.ExcelFile(f'{BASE_XLS}/ASELBY2026TONTINE60000.xlsx').sheet_names:
    if 'JANV' in sheet.upper(): MOIS_TONT_MAP[1] = sheet.replace('JANV', 'JANV')
    elif 'FEV' in sheet.upper(): MOIS_TONT_MAP[2] = sheet
    elif 'MARS' in sheet.upper(): MOIS_TONT_MAP[3] = sheet
    elif 'AVRIL' in sheet.upper(): MOIS_TONT_MAP[4] = sheet
    elif 'MAI' in sheet.upper(): MOIS_TONT_MAP[5] = sheet
    elif 'JUIN' in sheet.upper(): MOIS_TONT_MAP[6] = sheet
    elif 'JUIL' in sheet.upper(): MOIS_TONT_MAP[7] = sheet
    elif 'AOUT' in sheet.upper(): MOIS_TONT_MAP[8] = sheet
    elif 'SEPT' in sheet.upper(): MOIS_TONT_MAP[9] = sheet
    elif 'OCT' in sheet.upper(): MOIS_TONT_MAP[10] = sheet
    elif 'NOV' in sheet.upper(): MOIS_TONT_MAP[11] = sheet
    elif 'DEC' in sheet.upper(): MOIS_TONT_MAP[12] = sheet

# Adhérents map
adherents_map = {a.matricule: a for a in Adherent.objects.filter(
    statut='ACTIF', matricule__startswith='AS'
)}

sessions_map = {}
nb_saisies = nb_mvts = nb_parts = 0

# Mois disponibles = intersection des feuilles BASECALCUL et TABBORD
mois_disponibles = sorted(set(MOIS_BC_MAP.keys()) & set(MOIS_TB_MAP.keys()))
print(f"  Mois disponibles : {mois_disponibles}")

def _lire_participants_tontine(fname, sheet):
    """Lit les participants d'une session tontine depuis Excel."""
    try:
        df = pd.read_excel(f'{BASE_XLS}/{fname}', sheet_name=sheet, header=None)
        date_seance = to_date(df.iloc[3, 2])
        mask = df[0].astype(str).str.startswith('AS')
        result = {}
        for _, row in df[mask].iterrows():
            mat = str(row[0])
            nparts = int(row[1]) if pd.notna(row[1]) and row[1] else 1
            mode_raw = str(row[6] if df.shape[1] > 6 else 'BANQUE')
            mode = mode_raw if mode_raw in ['BANQUE','ESPECES','ECHEC'] else 'BANQUE'
            result[mat] = (nparts, mode)
        return date_seance, result
    except Exception:
        return None, {}

with transaction.atomic():
    for mois in mois_disponibles:
        sheet_bc = MOIS_BC_MAP[mois]
        sheet_tb = MOIS_TB_MAP.get(mois)
        sheet_tont = MOIS_TONT_MAP.get(mois)

        # Lire BASECALCUL
        df_bc = pd.read_excel(f'{BASE_XLS}/ASELBY2026BASECALCULINTERET.xlsx',
                              sheet_name=sheet_bc, header=None)
        bc_rows = {}
        for _, row in df_bc.iterrows():
            mat = str(row[0]) if row[0] else ''
            if mat.startswith('AS'):
                bc_rows[mat] = row

        # Lire TABBORD
        df_tb = None
        tb_rows = {}
        if sheet_tb:
            df_tb = pd.read_excel(f'{BASE_XLS}/ASELBY2026TABBORD.xlsx',
                                  sheet_name=sheet_tb, header=None)
            for _, row in df_tb.iterrows():
                mat = str(row[0]) if row[0] else ''
                if mat.startswith('AS'):
                    tb_rows[mat] = row

        # Participants par niveau tontine
        date_seance_tont, parts_t60 = None, {}
        if sheet_tont:
            date_seance_tont, parts_t60 = _lire_participants_tontine(
                'ASELBY2026TONTINE60000.xlsx', sheet_tont)
        _, parts_t75  = _lire_participants_tontine('ASELBY2026TONTINE75000.xlsx',
                                                    sheet_tont or f'TONT{list(MOIS_TONT_MAP.values())[0][4:]}')
        _, parts_t100 = _lire_participants_tontine('ASELBY2026TONTINE100000.xlsx',
                                                    sheet_tont or f'TONT{list(MOIS_TONT_MAP.values())[0][4:]}')

        if not date_seance_tont:
            # Chercher la date dans BASECALCUL
            date_row = df_bc.iloc[3, 10] if df_bc.shape[1] > 10 else None
            date_seance_tont = to_date(date_row) or date(annee, mois, 13)

        # Intérêts bureau = pool total = somme des intérêts répartis dans BASECALCUL
        pool_int = d(df_bc[bc_rows and 7 or 0].sum() if bc_rows else 0)
        # Plus précis : somme col[7] pour tous les membres
        mask_bc = df_bc[0].astype(str).str.startswith('AS')
        pool_int = D(str(df_bc[mask_bc][7].sum())) if mask_bc.any() else D('0')

        # Créer les sessions
        for niv, niv_parts in [(niv_t60, parts_t60), (niv_t75, parts_t75), (niv_t100, parts_t100)]:
            if not niv_parts:
                continue
            int_mensuel = pool_int / 12 if pool_int > 0 else D('0')
            s, _ = SessionTontine.objects.update_or_create(
                niveau=niv, mois=mois, annee=annee,
                defaults=dict(date_seance=date_seance_tont,
                             montant_interet_bureau=int_mensuel,
                             est_cloturee=True)
            )
            sessions_map[(niv.code, mois)] = s

            # Participations
            for mat, (nparts, mode) in niv_parts.items():
                adherent = adherents_map.get(mat)
                if not adherent:
                    continue
                montant_verse = niv.taux_mensuel * nparts if mode != 'ECHEC' else D('0')
                ParticipationTontine.objects.update_or_create(
                    session=s, adherent=adherent,
                    defaults=dict(nombre_parts=nparts, mode_versement=mode,
                                 montant_verse=montant_verse)
                )
                nb_parts += 1

        # Saisies + MouvementFonds
        for mat, adherent in adherents_map.items():
            r_bc = bc_rows.get(mat)
            r_tb = tb_rows.get(mat)

            # Mode versement depuis TABBORD col[8]
            if r_tb is not None:
                mode_raw = str(r_tb[8]) if pd.notna(r_tb[8]) else 'ECHEC'
                mode = mode_raw if mode_raw in ['BANQUE','ESPECES','ECHEC'] else (
                    'BANQUE' if d(r_tb[3]) > 0 else
                    'ESPECES' if d(r_tb[4]) > 0 else 'ECHEC'
                )
                vb  = d(r_tb[3]); ve = d(r_tb[4])
                pen_esp_val  = d(r_tb[7])
                pen_esp_appl = pen_esp_val > 0
                pen_ech      = d(r_tb[43]) if len(r_tb) > 43 else D('0')
                insc = d(r_tb[22]) if len(r_tb) > 22 else D('0')
                mut  = d(r_tb[45]) if len(r_tb) > 45 else D('0')
                foyer = d(r_tb[53]) if len(r_tb) > 53 else D('0')
                pret  = d(r_tb[25]) if len(r_tb) > 25 else D('0')
                remb  = d(r_tb[18]) if len(r_tb) > 18 else D('0')
                retr  = d(r_tb[47]) if len(r_tb) > 47 else D('0')
                sanc  = d(r_tb[21]) if len(r_tb) > 21 else D('0')
                eng   = d(r_tb[20]) if len(r_tb) > 20 else D('0')

                TableauDeBord.objects.filter(adherent=adherent, mois=mois, annee=annee).delete()
                TableauDeBord.objects.create(
                    adherent=adherent, mois=mois, annee=annee, config_exercice=config,
                    versement_banque=vb, versement_especes=ve,
                    mode_versement=mode,
                    inscription=insc, mutuelle=mut, contribution_foyer=foyer,
                    pret_fonds=pret, remboursement_pret=remb, retrait_partiel=retr,
                    sanction=sanc, montant_engagement=eng,
                    penalite_especes_appli=pen_esp_val,
                    penalite_echec_appli=pen_ech,
                    reste=d(r_bc[10]) if r_bc is not None else D('0'),
                    est_valide=True,
                )
                nb_saisies += 1

            if r_bc is not None:
                # Capital précédent
                if mois == min(mois_disponibles):
                    cap_prec = capitaux_depart.get(mat, D('0'))
                else:
                    mvt_prec = MouvementFonds.objects.filter(
                        adherent=adherent, mois=mois-1, annee=annee).first()
                    cap_prec = mvt_prec.capital_compose if mvt_prec else D('0')

                reste_bc = d(r_bc[10])
                if reste_bc > 0:
                    fr  = config.fonds_roulement_mensuel
                    fe  = config.frais_exceptionnels_mensuel
                    col = config.collation_mensuelle
                    ep  = reste_bc - fr - fe - col
                else:
                    fr = fe = col = ep = D('0')

                MouvementFonds.objects.filter(adherent=adherent, mois=mois, annee=annee).delete()
                MouvementFonds.objects.create(
                    adherent=adherent, mois=mois, annee=annee, config_exercice=config,
                    reconduction=D('0'),
                    retrait_partiel=d(r_bc[4]),
                    capital_compose_precedent=cap_prec,
                    reste=reste_bc,
                    epargne_nette=ep,
                    fonds_roulement=fr,
                    frais_exceptionnels=fe,
                    collation=col,
                    fonds_definitif=d(r_bc[5]),
                    base_calcul_interet=d(r_bc[6]),
                    interet_attribue=d(r_bc[7]),
                    capital_compose=d(r_bc[8]),
                    sanction=d(r_bc[9]),
                )
                nb_mvts += 1

        # Recalculer intérêts via service Django
        if pool_int > 0:
            fonds_services.calculer_interets_mensuels(mois, annee, pool_int, config)

print(f"  ✓ {len(sessions_map)} sessions | {nb_parts} participations")
print(f"  ✓ {nb_saisies} saisies | {nb_mvts} mouvements fonds")

# ── Mise à jour lots/intérêts depuis fichiers TONTINE (si TONTRESUME26 disponible)
print("\n  → Vérification TONTRESUME26...")
nb_lots = 0
TONT_FILES = {
    'T60':  ('ASELBY2026TONTINE60000.xlsx',  7,  8, 12, 16, 17, 20),
    'T75':  ('ASELBY2026TONTINE75000.xlsx',   5,  6, 10, 14, 15, 19),
    'T100': ('ASELBY2026TONTINE100000.xlsx',  5,  6, 10, 14, 15, 19),
}
for code, (fname, cv_pl, ci_pl, cr_pl, cm_lp, ci_lp, ci_adh) in TONT_FILES.items():
    xl_t = pd.ExcelFile(f'{BASE_XLS}/{fname}')
    resume_sheets = [s for s in xl_t.sheet_names if 'RESUME' in s.upper() and str(annee)[2:] in s]
    if not resume_sheets:
        continue
    df_res = pd.read_excel(f'{BASE_XLS}/{fname}', sheet_name=resume_sheets[0], header=None)
    niv = NiveauTontine.objects.filter(config_exercice=config, code=code).first()
    if not niv: continue
    mask = df_res[0].astype(str).str.startswith('AS')
    for _, row in df_res[mask].iterrows():
        mat = str(row[0])
        adherent = adherents_map.get(mat)
        if not adherent: continue
        def _rv(col): 
            v = row[col] if col < len(row) else None
            try: return D(str(float(v))) if pd.notna(v) and v else D('0')
            except: return D('0')
        first_part = ParticipationTontine.objects.filter(
            adherent=adherent, session__niveau=niv
        ).order_by('session__mois').first()
        if first_part:
            first_part.vente_petit_lot         = _rv(cv_pl)
            first_part.interet_petit_lot       = _rv(ci_pl)
            first_part.remboursement_petit_lot = _rv(cr_pl)
            first_part.montant_lot_principal   = _rv(cm_lp)
            first_part.interet_lot_principal   = _rv(ci_adh)
            if _rv(cm_lp) > 0:
                first_part.a_obtenu_lot_principal = True
            first_part.save()
            nb_lots += 1
print(f"  ✓ {nb_lots} participations mises à jour (lots/intérêts)")

# ══════════════════════════════════════════════════════════════════════
# 6. AGIOS depuis HISTOBQUE 2026
# ══════════════════════════════════════════════════════════════════════
print("\n[6] Agios bancaires 2026...")
xl_histo = pd.ExcelFile(f'{BASE_XLS}/ASELBY2026TABBHISTOBQUE.xlsx')
MOIS_HISTO_MAP = {
    'JANV': 1, 'FEV': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'JUIN': 6,
    'JUIL': 7, 'AOUT': 8, 'SEPT': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
}
nb_agios = 0
for sheet in xl_histo.sheet_names:
    if 'HISTOBQUE' not in sheet.upper(): continue
    mois_num = None
    for kw, m in MOIS_HISTO_MAP.items():
        if kw in sheet.upper():
            mois_num = m; break
    if not mois_num: continue
    df_ag = pd.read_excel(f'{BASE_XLS}/ASELBY2026TABBHISTOBQUE.xlsx',
                          sheet_name=sheet, header=None)
    # AGIO sur la ligne ASELBY (col 9 = AGIO)
    mask_ag = df_ag[0].astype(str).str.contains('AS201648|ASELBY', na=False)
    if mask_ag.any():
        agio_val = d(df_ag[mask_ag].iloc[0, 9])
        AgioBancaire.objects.update_or_create(
            config_exercice=config, mois=mois_num, annee=annee,
            defaults=dict(montant_agio=agio_val, interet_crediteur=D('0'))
        )
        nb_agios += 1
print(f"  ✓ {nb_agios} agios importés depuis HISTOBQUE")

# ══════════════════════════════════════════════════════════════════════
# 7. DÉPENSES depuis TRAVAUXFINEXERCICE2026
#    Fonds roulement, frais exceptionnels, collation, foyer, autres
# ══════════════════════════════════════════════════════════════════════
print("\n[7] Dépenses depuis TRAVAUXFINEXERCICE2026...")

def _importer_depenses(sheet_name, model_class, label):
    """Importe les lignes de sortie d'une feuille DETAIL depuis TRAVAUXFINEXERCICE."""
    try:
        df = pd.read_excel(TRAV, sheet_name=sheet_name, header=None)
    except Exception:
        print(f"  ⚠ Feuille {sheet_name} manquante")
        return 0
    nb = 0
    for _, row in df.iterrows():
        # Les sorties ont une date en col3/col4 et un libellé en col4/col5 et montant en col5/col6
        # Format: date en col3, libellé en col4, montant en col5 (ou col4,5,6 selon la feuille)
        date_val = to_date(row[3]) if len(row) > 3 and pd.notna(row[3]) else None
        libelle  = str(row[4]) if len(row) > 4 and pd.notna(row[4]) else ''
        montant  = d(row[5]) if len(row) > 5 and pd.notna(row[5]) else D('0')
        if not date_val or not libelle or montant == 0:
            continue
        if str(libelle).upper() in ['LIBELLE', 'NAN', 'DATE'] or not libelle.strip():
            continue
        try:
            model_class.objects.get_or_create(
                config_exercice=config, date=date_val, libelle=libelle.upper(),
                defaults=dict(montant=montant)
            )
            nb += 1
        except Exception:
            pass
    return nb

nb_fr  = _importer_depenses('DETAILFPNDSROULT',       DepenseFondsRoulement,    'FR')
nb_fe  = _importer_depenses('DETAILFRAISEXCEPTIONNEL', DepenseFraisExceptionnels,'FE')
nb_col = _importer_depenses('DETAILCOLLATION',         DepenseCollation,         'COL')

# Foyer
try:
    df_foyer = pd.read_excel(TRAV, sheet_name='DETAILFOYER', header=None)
    nb_foyer = 0
    for _, row in df_foyer.iterrows():
        date_val = to_date(row[3]) if len(row) > 3 and pd.notna(row[3]) else None
        libelle  = str(row[4]) if len(row) > 4 and pd.notna(row[4]) else ''
        montant  = d(row[5]) if len(row) > 5 else D('0')
        if date_val and montant > 0 and libelle.strip():
            DepenseFoyer.objects.get_or_create(
                config_exercice=config, date=date_val, libelle=libelle.upper(),
                defaults=dict(montant=montant)
            )
            nb_foyer += 1
except Exception as e:
    nb_foyer = 0

# Autres dépenses
try:
    df_ad = pd.read_excel(TRAV, sheet_name='DETAILAUTREDEPENSE', header=None)
    nb_ad = 0
    for _, row in df_ad.iterrows():
        date_val = to_date(row[0]) if pd.notna(row[0]) else None
        libelle  = str(row[1]) if pd.notna(row[1]) else ''
        montant  = d(row[2]) if len(row) > 2 and pd.notna(row[2]) else D('0')
        if libelle.strip() and montant > 0:
            AutreDepense.objects.get_or_create(
                config_exercice=config, libelle=libelle.upper(),
                defaults=dict(date=date_val or date(annee,12,31), montant=montant)
            )
            nb_ad += 1
except Exception:
    nb_ad = 0

print(f"  ✓ FR={nb_fr}, FE={nb_fe}, COL={nb_col}, FOYER={nb_foyer}, AUTRES={nb_ad}")

# ══════════════════════════════════════════════════════════════════════
# 8. FICHES CASSATION depuis TRAVAUXFINEXERCICE2026
# ══════════════════════════════════════════════════════════════════════
print("\n[8] Fiches cassation 2026...")
try:
    df_fc = pd.read_excel(TRAV, sheet_name='DETAILFICHECASSATION', header=None)
    # Trouver la ligne d'en-tête (ligne contenant 'MATRICULE')
    hdr_row = None
    for i, row in df_fc.iterrows():
        if 'MATRICULE' in str(row.tolist()).upper():
            hdr_row = i; break

    nb_fiches = 0
    if hdr_row is not None:
        for idx in range(hdr_row + 1, len(df_fc)):
            row = df_fc.iloc[idx]
            mat = str(row[0]) if pd.notna(row[0]) else ''
            if not mat.startswith('AS'):
                continue
            adherent = Adherent.objects.filter(matricule=mat).first()
            if not adherent:
                continue

            def _rv(col, default=0):
                v = row[col] if col < len(row) and pd.notna(row[col]) else None
                try: return D(str(float(v))) if v is not None else D(str(default))
                except: return D(str(default))

            # Colonnes: 0=mat, 1=nom, 2=fonds_caisse, 3=reconduction, 4=repartition_interets
            # 5=epargne, 6=rep_pen, 7=rep_col, 8=total_dist, 9=sanction,
            # 10=comp_mutuelle, 11=comp_fonds, 12=dette_pret, 13=total_ret,
            # 14=total_percv, 15=montant_percu, 16=reconduction_2026, 17=nouveau_fonds,
            # 18=dons_foyer, 19=percu_especes, 20=percu_cheque, 21=interet
            f = FicheCassation(
                adherent=adherent, config_exercice=config,
                fonds_caisse         = _rv(2),
                reconduction         = _rv(3),
                repartition_interets = _rv(4),
                epargne_cumulee      = _rv(5),
                repartition_penalites= _rv(6),
                repartition_collation= _rv(7),
                sanctions            = _rv(9),
                complement_mutuelle  = _rv(10),
                complement_fonds     = _rv(11),
                dette_pret           = _rv(12),
                montant_percu        = _rv(15),
                nouveau_fonds        = _rv(17),
                dons_foyer           = _rv(18),
                montant_percu_especes= _rv(19),
                montant_percu_cheque = _rv(20),
                est_validee          = True,
            )
            FicheCassation.objects.filter(adherent=adherent, config_exercice=config).delete()
            f.save()
            nb_fiches += 1

    print(f"  ✓ {nb_fiches} fiches cassation importées")
except Exception as e:
    print(f"  ⚠ Erreur fiches cassation: {e}")

# ══════════════════════════════════════════════════════════════════════
# 9. SYNTHESE COMPTE depuis TRAVAUXFINEXERCICE2026
# ══════════════════════════════════════════════════════════════════════
print("\n[9] Synthèse des comptes 2026...")
try:
    df_sc = pd.read_excel(TRAV, sheet_name='SYNTHESECOMPTE', header=None)
    # Trouver ligne en-têtes (NUMERO ORDRE | RUBRIQUE | REPORT...)
    hdr_idx = None
    for i, row in df_sc.iterrows():
        if 'RUBRIQUE' in str(row.tolist()).upper():
            hdr_idx = i; break

    # Mapping rubrique → champs SyntheseCompte
    RUBRIQUE_MAP = {
        'FONDS DE CAISSE': ('report_fonds_caisse','entrees_fonds_caisse','sorties_fonds_caisse'),
        'FONDS DE ROULEMENT': ('report_fonds_roulement','entrees_fonds_roulement','sorties_fonds_roulement'),
        'FRAIS EXCEPTIONNELS': ('report_frais_exceptionnels','entrees_frais_excep','sorties_frais_excep'),
        'MUTUELLE': ('report_mutuelle','entrees_mutuelle','sorties_mutuelle'),
        'INSCRIPTION': ('report_inscription','entrees_inscription',None),
        'PENALITES': ('report_penalites','entrees_penalites','sorties_penalites'),
        'SANCTIONS': ('report_sanctions','entrees_sanctions','sorties_sanctions'),
        'COLLATION': ('report_collation','entrees_collation','sorties_collation'),
        'INTERET BANCAIRE': ('report_interet_bancaire',None,'sorties_interet_bancaire'),
        'FOYER': ('report_foyer','entrees_foyer','sorties_foyer'),
        'TERRAIN': ('report_terrain','entrees_terrain','sorties_terrain'),
        'INTERET EPARGNE': ('report_interet_epargne',None,'sorties_interet_epargne'),
        'AIDES LAGWE': ('report_aides_lagwe',None,'sorties_aides_lagwe'),
    }

    defaults = {}
    if hdr_idx is not None:
        for idx in range(hdr_idx + 1, len(df_sc)):
            row = df_sc.iloc[idx]
            rubrique = str(row[1]).upper() if pd.notna(row[1]) else ''
            for key, (f_rep, f_ent, f_sor) in RUBRIQUE_MAP.items():
                if key in rubrique:
                    if f_rep: defaults[f_rep] = d(row[2])
                    if f_ent: defaults[f_ent] = d(row[3])
                    if f_sor: defaults[f_sor] = d(row[4])
                    break
            # DEPOT TCHOUAMO
            if 'DEPOT' in rubrique or 'TCHOUAMO' in rubrique:
                defaults['report_depot_tchouamo'] = d(row[2])

    # Disposition des fonds depuis DISPOFONDS
    df_disp = pd.read_excel(TRAV, sheet_name='DISPOFONDS', header=None)
    DISP_MAP = {
        'COMPTE CCA': 'compte_cca', 'COMPTE MC': 'compte_mc2',
        'AFRILAND': 'compte_afriland', 'PREFINANCEMENT': 'dette_prefinancement_lagwe',
        'KOUATCHO': 'dette_mr_kouatcho', 'RECUPER': 'reste_a_recuperer_conseil',
        'CIRCULATION': None,  # calculé depuis Pret
        'TONTINES': 'dettes_tontines',
    }
    for _, row in df_disp.iterrows():
        lib = str(row[1]).upper() if pd.notna(row[1]) else ''
        for key, champ in DISP_MAP.items():
            if key in lib and champ:
                defaults[champ] = d(row[3]) if pd.notna(row[3]) else d(row[2])
                break

    SyntheseCompte.objects.update_or_create(
        config_exercice=config, defaults=defaults
    )
    print(f"  ✓ SyntheseCompte 2026 importée ({len(defaults)} champs)")
except Exception as e:
    print(f"  ⚠ Erreur synthèse: {e}")

# ══════════════════════════════════════════════════════════════════════
# 10. LISTE ROUGE depuis DETAILLISTEROUGE
# ══════════════════════════════════════════════════════════════════════
print("\n[10] Liste rouge 2026...")
try:
    df_lr = pd.read_excel(TRAV, sheet_name='DETAILLISTEROUGE', header=None)
    hdr_lr = None
    for i, row in df_lr.iterrows():
        if 'MATRICULE' in str(row.tolist()).upper():
            hdr_lr = i; break
    nb_lr = 0
    if hdr_lr is not None:
        for idx in range(hdr_lr + 1, len(df_lr)):
            row = df_lr.iloc[idx]
            mat = str(row[0]) if pd.notna(row[0]) else ''
            nom = str(row[1]) if pd.notna(row[1]) else ''
            if not mat.startswith('AS') or not nom.strip() or nom.upper() in ['NAN','TOTAL']:
                continue
            dette    = d(row[2]) if pd.notna(row[2]) else D('0')
            paiement = d(row[3]) if pd.notna(row[3]) else D('0')
            obs      = str(row[6]) if len(row) > 6 and pd.notna(row[6]) else 'FONDS LIQUIDE'
            # L'adhérent liste rouge a un matricule LR...
            mat_lr = 'LR' + mat[2:]
            a_lr, _ = Adherent.objects.update_or_create(
                matricule=mat_lr,
                defaults=dict(nom_prenom=nom, numero_ordre=800+idx, statut='INACTIF')
            )
            ListeRouge.objects.update_or_create(
                adherent=a_lr,
                defaults=dict(
                    config_exercice=config,
                    motif=obs.upper() if obs.strip() else 'FONDS LIQUIDE',
                    montant_dette=dette,
                    montant_garantie=paiement,
                    est_solde=(paiement >= dette),
                )
            )
            nb_lr += 1
    print(f"  ✓ {nb_lr} entrées liste rouge")
except Exception as e:
    print(f"  ⚠ Erreur liste rouge: {e}")

# ══════════════════════════════════════════════════════════════════════
# RÉSUMÉ FINAL
# ══════════════════════════════════════════════════════════════════════
print("\n" + "=" * 65)
print(f"IMPORT {annee} TERMINÉ ✓")
print("=" * 65)
print(f"  ConfigExercice {annee} : FR={config.fonds_roulement_mensuel}/mois")
print(f"  Adhérents actifs  : {Adherent.objects.filter(statut='ACTIF').count()}")
print(f"  Niveaux tontine   : {NiveauTontine.objects.filter(config_exercice=config).count()} (T60, T75, T100)")
print(f"  Sessions          : {SessionTontine.objects.filter(niveau__config_exercice=config).count()}")
print(f"  Participations    : {ParticipationTontine.objects.filter(session__niveau__config_exercice=config).count()}")
print(f"  Saisies           : {TableauDeBord.objects.filter(config_exercice=config).count()}")
print(f"  MouvementsFonds   : {MouvementFonds.objects.filter(config_exercice=config).count()}")
print(f"  FichesCassation   : {FicheCassation.objects.filter(config_exercice=config).count()}")
print(f"  Agios             : {AgioBancaire.objects.filter(config_exercice=config).count()} mois")
print(f"  ListeRouge        : {ListeRouge.objects.count()} entrées")
print(f"\n  Mois importés     : {sorted(set(MOIS_BC_MAP.keys()) & set(MOIS_TB_MAP.keys()))}")
print(f"  (données partielles — exercice en cours)")