# Services saisie
"""
apps/saisie/services.py
Logique de calcul centralisée.
Appelée après chaque save de ReleveBancaire ou SaisieMonthly.
"""
from decimal import Decimal
from django.db import transaction
D = Decimal


def calculer_mouvement_fonds(saisie):
    """
    Crée ou met à jour MouvementFonds depuis une SaisieMonthly.
    Récupère le capital_compose du mois précédent automatiquement.
    """
    from apps.fonds.models import MouvementFonds

    adherent = saisie.adherent
    mois     = saisie.mois
    annee    = saisie.annee
    config   = saisie.config_exercice

    # Capital de départ = capital_compose du mois précédent
    prec_m = 12 if mois == 1 else mois - 1
    prec_a = annee - 1 if mois == 1 else annee
    mvt_prec = MouvementFonds.objects.filter(
        adherent=adherent, mois=prec_m, annee=prec_a
    ).first()
    capital_prec = mvt_prec.capital_compose if mvt_prec else D('0')

    # Créer ou récupérer
    mvt, _ = MouvementFonds.objects.get_or_create(
        adherent=adherent, mois=mois, annee=annee,
        defaults={'config_exercice': config,
                  'capital_compose_precedent': capital_prec}
    )
    mvt.config_exercice            = config
    mvt.capital_compose_precedent  = capital_prec
    mvt.calculer_depuis_saisie(saisie)
    mvt.save()
    return mvt


def repartir_interets(mois, annee, pool_interets, config):
    """
    Répartit le pool d'intérêts entre tous les adhérents éligibles.
    Col H = ROUNDDOWN(pool × base_i / total_bases, 2)
    """
    from apps.fonds.models import MouvementFonds, ReserveMensuelle

    mvts = list(MouvementFonds.objects.filter(
        mois=mois, annee=annee, config_exercice=config
    ))

    total_bases = sum(m.base_calcul_interet for m in mvts
                      if m.base_calcul_interet > 0)
    nb_eligibles = sum(1 for m in mvts if m.base_calcul_interet > 0)

    reserve, _ = ReserveMensuelle.objects.update_or_create(
        mois=mois, annee=annee, config_exercice=config,
        defaults={
            'pool_interets':          D(str(pool_interets)),
            'total_bases_eligibles':  total_bases,
            'nb_adherents_eligibles': nb_eligibles,
        }
    )

    for mvt in mvts:
        mvt.appliquer_interet(pool_interets, total_bases)

    MouvementFonds.objects.bulk_update(
        mvts, ['interet_attribue', 'capital_compose'])

    reserve.est_reparti = True
    reserve.save()
    return reserve


def recalculer_tout_le_mois(mois, annee, config):
    """
    Recalcule tous les MouvementFonds d'un mois.
    Utile après correction d'une saisie.
    """
    from apps.saisie.models import SaisieMonthly
    saisies = SaisieMonthly.objects.filter(
        mois=mois, annee=annee, config_exercice=config
    ).select_related('adherent')
    with transaction.atomic():
        for s in saisies:
            calculer_mouvement_fonds(s)