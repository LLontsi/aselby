from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect
from django.utils import timezone
from django.db.models import Sum
from decimal import Decimal

from apps.parametrage.models import ConfigExercice
from apps.adherents.models import Adherent
from apps.tontines.models import NiveauTontine, ParticipationTontine, SessionTontine
from apps.fonds.models import MouvementFonds
from apps.prets.models import Pret
from apps.saisie.models import TableauDeBord
from apps.dettes.models import ListeRouge

MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre']

def _bureau_required(view_func):
    @login_required
    def wrapper(request, *args, **kwargs):
        if not request.user.est_bureau:
            return redirect('users:mon_espace')
        return view_func(request, *args, **kwargs)
    return wrapper

@_bureau_required
def dashboard(request):
    config = ConfigExercice.get_exercice_courant()
    mois   = timezone.now().month
    annee  = timezone.now().year

    nb_actifs = Adherent.objects.filter(statut='ACTIF').count()

    # Chercher les données du mois courant ou dernier mois disponible
    saisies_mois = TableauDeBord.objects.filter(mois=mois, annee=annee, config_exercice=config)
    if not saisies_mois.exists():
        dernier_mvt = TableauDeBord.objects.filter(
            config_exercice=config
        ).order_by('-annee', '-mois').first()
        if dernier_mvt:
            mois  = dernier_mvt.mois
            annee = dernier_mvt.annee
            saisies_mois = TableauDeBord.objects.filter(mois=mois, annee=annee, config_exercice=config)

    total_fonds = MouvementFonds.objects.filter(
        annee=annee, mois=mois, config_exercice=config
    ).aggregate(t=Sum('fonds_definitif'))['t'] or Decimal('0')

    # Prêts : tous les prêts EN_COURS (pas de filtre config)
    prets_qs    = Pret.objects.filter(statut=Pret.EN_COURS)
    total_prets = prets_qs.aggregate(t=Sum('montant_total_du'))['t'] or Decimal('0')

    demandes_attente = Pret.objects.filter(est_demande_membre=True, est_valide_bureau=False)
    prets_retard     = prets_qs.filter(nb_mois_retard__gt=0)

    alertes = []
    if prets_retard.exists():
        alertes.append({'type': 'retard', 'message': f"{prets_retard.count()} adhérent(s) n'ont pas remboursé leur prêt ce mois"})
    if demandes_attente.exists():
        alertes.append({'type': 'info', 'message': f"{demandes_attente.count()} demande(s) de prêt en attente de validation"})
    nb_saisis = saisies_mois.count()
    if nb_saisis < nb_actifs:
        alertes.append({'type': 'info', 'message': f"Saisie : {nb_saisis}/{nb_actifs} adhérents saisis pour {MOIS_FR[mois]}"})
    else:
        alertes.append({'type': 'succes', 'message': f"Saisie complète pour {MOIS_FR[mois]}"})

    saisies_recentes = saisies_mois.select_related('adherent').order_by('adherent__numero_ordre')[:10]

    # Participations tontines du mois courant
    participations_tontines = []
    for niveau in NiveauTontine.objects.filter(config_exercice=config).order_by('taux_mensuel'):
        sess = SessionTontine.objects.filter(niveau=niveau, mois=mois, annee=annee).first()
        nb_adh   = ParticipationTontine.objects.filter(session=sess).values('adherent').distinct().count() if sess else 0
        nb_parts = ParticipationTontine.objects.filter(session=sess).aggregate(t=Sum('nombre_parts'))['t'] or 0 if sess else 0
        participations_tontines.append((niveau.code, nb_parts, nb_adh))

    ctx = {
        'config_exercice'      : config,
        # Variables attendues par le template dashboard.html
        'nb_adherents'         : nb_actifs,
        'nb_saisis'            : nb_saisis,
        'nb_total'             : nb_actifs,
        'prets_en_cours'       : prets_qs.count(),
        'nb_liste_rouge'       : ListeRouge.objects.filter(est_solde=False).count(),
        'mois_courant'         : f"{MOIS_FR[mois]} {annee}",
        'entrees_prevues'      : total_fonds,
        'depenses_prevues'     : Decimal('0'),
        'solde_disponible'     : total_fonds,
        'alertes'              : alertes,
        'dernieres_saisies'    : saisies_recentes,
        'participations_tontines': participations_tontines,
        'prets_retard'         : prets_retard.select_related('adherent').order_by('-nb_mois_retard')[:10],
        'demandes_prets'       : demandes_attente.select_related('adherent')[:5],
        # Garder aussi les anciennes clés pour la compatibilité
        'kpi': {
            'nb_adherents_actifs'    : nb_actifs,
            'total_prets_circulation': total_prets,
            'nb_prets_en_cours'      : prets_qs.count(),
            'nb_liste_rouge'         : ListeRouge.objects.filter(est_solde=False).count(),
        },
    }
    return render(request, 'dashboard/dashboard.html', ctx)