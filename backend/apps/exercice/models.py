from django.db import models
from decimal import Decimal


class FicheCassation(models.Model):
    """Fiche de cassation annuelle par adhérent."""
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='fiches_cassation')
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)

    # Ce qui est dû
    fonds_caisse              = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    repartition_interets      = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    epargne_cumulee           = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    repartition_penalites     = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    repartition_collation     = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # Retenues
    sanctions                 = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    complement_mutuelle       = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    complement_fonds          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    dette_pret                = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # Distribution
    dons_foyer                = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    montant_percu             = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    montant_percu_especes     = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    montant_percu_cheque      = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # Résultats calculés
    reconduction              = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    nouveau_fonds             = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    est_validee = models.BooleanField(default=False)
    date_calcul = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Fiche de cassation"
        unique_together = [('adherent', 'config_exercice')]

    def __str__(self):
        return f"Cassation {self.adherent.matricule} - {self.config_exercice.annee}"

    @property
    def total_a_distribuer(self):
        """Excel col I: reconduction_reportée + interets + epargne + penalites_part + collation_part
        NB: fonds_caisse (col C) est affiché séparément et N'EST PAS inclus dans ce total."""
        return (
            self.reconduction
            + self.repartition_interets
            + self.epargne_cumulee
            + self.repartition_penalites
            + self.repartition_collation
        )

    @property
    def total_retenu(self):
        """Excel col N: comp_mutuelle + comp_fonds UNIQUEMENT.
        NB: sanctions (col J) et dette_pret (col M) sont affichés séparément."""
        return self.complement_mutuelle + self.complement_fonds

    @property
    def net_a_percevoir(self):
        """Excel col O: total_a_distribuer - total_retenu"""
        return self.total_a_distribuer - self.total_retenu

    def calculer_reconduction(self):
        """Excel col Q (RECONDUCTION prochaine année): net - montant_percu - dons_foyer.
        ATTENTION : n'écrase PAS self.reconduction (qui = solde reporté col D).
        Retourne seulement la valeur calculée sans modifier le champ."""
        return self.net_a_percevoir - self.montant_percu - self.dons_foyer


class SyntheseCompte(models.Model):
    """
    Bilan comptable annuel de l'association.
    17 rubriques identiques au fichier Excel TRAVAUXFINEXERCICE.
    """
    config_exercice = models.OneToOneField('parametrage.ConfigExercice', on_delete=models.PROTECT)

    # 1. Fonds de caisse et reconduction
    report_fonds_caisse        = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    entrees_fonds_caisse       = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    sorties_fonds_caisse       = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # 2. Fonds de roulement
    report_fonds_roulement     = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_fonds_roulement    = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_fonds_roulement    = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 3. Frais exceptionnels
    report_frais_exceptionnels = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_frais_excep        = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_frais_excep        = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 4. Mutuelle / Complément mutuel
    report_mutuelle            = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_mutuelle           = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_mutuelle           = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 5. Inscription
    report_inscription         = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_inscription        = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 6. Pénalités paiement espèces
    report_penalites           = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_penalites          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_penalites          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 7. Sanctions (NOUVEAU — manquait avant)
    report_sanctions           = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_sanctions          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_sanctions          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 8. Collation
    report_collation           = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_collation          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_collation          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 9. Intérêt bancaire (agios en sortie, intérêts créditeurs en report) (NOUVEAU)
    report_interet_bancaire    = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_interet_bancaire   = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 10. Foyer
    report_foyer               = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    entrees_foyer              = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_foyer              = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 11. Terrain (NOUVEAU — dépenses patrimoniales)
    report_terrain             = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    entrees_terrain            = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    sorties_terrain            = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # 12. Intérêt épargne (intérêts répartis sur les fonds) (NOUVEAU)
    report_interet_epargne     = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    sorties_interet_epargne    = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # 13. Aides LAGWE non reversées (NOUVEAU)
    report_aides_lagwe         = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sorties_aides_lagwe        = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # 14. Dépôt TCHOUAMO (NOUVEAU)
    report_depot_tchouamo      = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    sorties_depot_tchouamo     = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # ── DISPOSITION DES FONDS (comptes bancaires réels) ─────────────
    # Valeurs saisies par l'admin depuis les relevés bancaires
    compte_cca                 = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Compte CCA", help_text="Solde compte CCA")
    compte_mc2                 = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Compte MC² (non actualisé)")
    compte_afriland            = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Compte AFRILAND")
    dette_prefinancement_lagwe = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Dette préfinancement LA'AGWEU")
    dette_mr_kouatcho          = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Dette Mr KOUATCHO")
    reste_a_recuperer_conseil  = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Reste à récupérer au conseil")
    dettes_tontines            = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Dettes tontines")
    autres_disponibilites      = models.TextField(blank=True, default='',
        verbose_name="Autres disponibilités (libellé + montant)")

    date_calcul = models.DateTimeField(auto_now=True)

    def solde(self, report, entrees, sorties):
        return report + entrees - sorties

    @property
    def solde_fonds_caisse(self):
        return self.report_fonds_caisse + self.entrees_fonds_caisse - self.sorties_fonds_caisse

    @property
    def solde_fonds_roulement(self):
        return self.report_fonds_roulement + self.entrees_fonds_roulement - self.sorties_fonds_roulement

    @property
    def solde_collation(self):
        return self.report_collation + self.entrees_collation - self.sorties_collation

    @property
    def solde_foyer(self):
        return self.report_foyer + self.entrees_foyer - self.sorties_foyer