"""
Service central de calcul des intérêts sur fonds.
Fidèle aux formules Excel BASECALCULINTERET.
"""
from decimal import Decimal
from .models import MouvementFonds, ReserveMensuelle


def calculer_interets_mensuels(mois, annee, pool_interets, config_exercice):
    """
    Répartit le pool d'intérêts entre tous les adhérents éligibles.

    Excel:
      - Base individuelle G: =IF(fonds_def > seuil, capital_précédent - retrait + épargne, 0)
      - Répartition H: =ROUNDDOWN(pool_total / total_bases * base_i, 2)

    pool_interets: montant total décidé (pour fonds) ou calculé (petits lots)
    """
    mouvements = MouvementFonds.objects.filter(
        mois=mois, annee=annee, config_exercice=config_exercice
    ).select_related('adherent')

    # Calculer toutes les bases
    for mvt in mouvements:
        mvt.calculer_base_interet()

    # Total bases éligibles
    total_bases = sum(m.base_calcul_interet for m in mouvements if m.base_calcul_interet > 0)
    nb_eligibles = sum(1 for m in mouvements if m.base_calcul_interet > 0)

    # Créer ou mettre à jour la réserve mensuelle
    reserve, _ = ReserveMensuelle.objects.update_or_create(
        mois=mois, annee=annee,
        defaults={
            'config_exercice': config_exercice,
            'pool_interets': Decimal(str(pool_interets)),
            'total_bases_eligibles': total_bases,
            'nb_adherents_eligibles': nb_eligibles,
        }
    )

    # Répartir proportionnellement
    total_distribue = Decimal('0')
    bulk_update = []

    for mvt in mouvements:
        if mvt.base_calcul_interet > 0 and total_bases > 0:
            mvt.interet_attribue = reserve.calculer_interet_adherent(mvt.base_calcul_interet)
        else:
            mvt.interet_attribue = Decimal('0')
        mvt.calculer_capital_compose()
        total_distribue += mvt.interet_attribue
        bulk_update.append(mvt)

    MouvementFonds.objects.bulk_update(
        bulk_update,
        ['base_calcul_interet', 'interet_attribue', 'capital_compose']
    )

    reserve.est_reparti = True
    reserve.save()

    return {
        'pool': pool_interets,
        'total_bases': total_bases,
        'nb_eligibles': nb_eligibles,
        'total_distribue': total_distribue,
        'mouvements': bulk_update,
    }


def calculer_reste_mensuel(montant_tontine_verse, complement_epargne, taux_tontine_pur, mode):
    """
    Excel col AS: =IF(I="ECHEC", 0, (K + G - taux_pur))
    K = montant_tontine_verse (col K = nbre_parts * versement_par_part)
    G = complement_epargne
    taux_pur = taux de tontine pur (35000 pour T35)
    """
    if mode == 'ECHEC':
        return Decimal('0')
    return Decimal(str(montant_tontine_verse)) + Decimal(str(complement_epargne)) - Decimal(str(taux_tontine_pur))
