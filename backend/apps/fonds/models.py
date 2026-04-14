from django.db import models
from decimal import Decimal
import math


class MouvementFonds(models.Model):
    """
    Mouvement mensuel du fonds de caisse d'un adhérent.
    Toute la logique de calcul est ici, fidèle aux formules Excel.
    """
    adherent         = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='mouvements_fonds')
    mois             = models.IntegerField()
    annee            = models.IntegerField()
    config_exercice  = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)

    # Entrées
    reconduction           = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    retrait_partiel        = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    capital_compose_precedent = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # Calculés automatiquement (sauvegardés pour historique)
    reste                  = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    epargne_nette          = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    fonds_roulement        = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    frais_exceptionnels    = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    collation              = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    fonds_definitif        = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    base_calcul_interet    = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    interet_attribue       = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    capital_compose        = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    sanction               = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    date_calcul = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Mouvement fonds"
        verbose_name_plural = "Mouvements fonds"
        unique_together = [('adherent', 'mois', 'annee')]
        ordering = ['annee', 'mois']

    def __str__(self):
        return f"{self.adherent.matricule} - {self.mois:02d}/{self.annee}"

    # ----------------------------------------------------------------
    # METHODES DE CALCUL - Fidèles aux formules Excel
    # ----------------------------------------------------------------

    def calculer_epargne_nette(self):
        """
        Excel col L: =IF(K7=0, 0, K7-(M7+N7+O7))
        K = reste, M = fonds_roulement, N = frais_excep, O = collation
        """
        cfg = self.config_exercice
        if self.reste == 0:
            self.fonds_roulement     = Decimal('0')
            self.frais_exceptionnels = Decimal('0')
            self.collation           = Decimal('0')
            self.epargne_nette       = Decimal('0')
        else:
            self.fonds_roulement     = cfg.fonds_roulement_mensuel
            self.frais_exceptionnels = cfg.frais_exceptionnels_mensuel
            self.collation           = cfg.collation_mensuelle
            self.epargne_nette       = self.reste - cfg.charges_fixes_mensuelles
        return self.epargne_nette

    def calculer_fonds_definitif(self):
        """
        Excel col F: capital_composé_précédent + reconduction - retrait_partiel + épargne_nette
        """
        self.fonds_definitif = (
            self.capital_compose_precedent
            + self.reconduction
            - self.retrait_partiel
            + self.epargne_nette
        )
        return self.fonds_definitif

    def calculer_base_interet(self):
        """
        Excel col G: =IF(fonds_definitif > 900000, (capital_precedent - retrait + epargne), 0)
        Seuil issu de la config exercice.
        """
        seuil = self.config_exercice.seuil_eligibilite_interets
        if self.fonds_definitif > seuil:
            self.base_calcul_interet = (
                self.capital_compose_precedent
                - self.retrait_partiel
                + self.epargne_nette
            )
        else:
            self.base_calcul_interet = Decimal('0')
        return self.base_calcul_interet

    def calculer_capital_compose(self):
        """
        Excel col I: fonds_definitif + interet_attribue
        """
        self.capital_compose = self.fonds_definitif + self.interet_attribue
        return self.capital_compose

    def recalculer_tout(self, reste):
        """
        Recalcule dans l'ordre exact des formules Excel.
        Appelé après la saisie mensuelle.
        """
        self.reste = Decimal(str(reste))
        self.calculer_epargne_nette()
        self.calculer_fonds_definitif()
        self.calculer_base_interet()
        # interet_attribue est calculé globalement (service) puis injecté ici
        self.calculer_capital_compose()

    @property
    def est_eligible_interets(self):
        return self.base_calcul_interet > 0

    @property
    def taux_repartition(self):
        """Pourcentage de la base de cet adhérent sur le total éligible."""
        return self.base_calcul_interet


class ReserveMensuelle(models.Model):
    """
    Pool mensuel d'intérêts à répartir entre les adhérents éligibles.
    Calculé une fois par mois, puis réparti.
    """
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    pool_interets   = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    total_bases_eligibles = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    nb_adherents_eligibles = models.IntegerField(default=0)
    est_reparti     = models.BooleanField(default=False)

    class Meta:
        unique_together = [('mois', 'annee')]
        ordering = ['annee', 'mois']

    def __str__(self):
        return f"Réserve {self.mois:02d}/{self.annee} - {self.pool_interets:,.2f} FCFA"

    def calculer_interet_adherent(self, base_individuelle):
        """
        Excel: =ROUNDDOWN(pool_total / total_bases * base_i, 2)
        """
        if self.total_bases_eligibles == 0:
            return Decimal('0')
        ratio = float(self.pool_interets) / float(self.total_bases_eligibles)
        interet = ratio * float(base_individuelle)
        # ROUNDDOWN à 2 décimales = floor(x * 100) / 100
        return Decimal(str(math.floor(interet * 100) / 100))
