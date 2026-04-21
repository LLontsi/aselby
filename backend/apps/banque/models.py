
from django.db import models
from decimal import Decimal


class HistoriqueBancaire(models.Model):
    adherent           = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT,
                                           related_name='historique_bancaire')
    mois               = models.IntegerField()
    annee              = models.IntegerField()
    versement_tontine  = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    versement_especes  = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    versement_banque   = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    autre_versement    = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    montant_engagement = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    agio               = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    # Nouveaux champs pour TABBHISTOBQUE
    en_compte_reel          = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="En compte", help_text="Montant constaté en compte")
    montant_a_justifier_saisi = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'),
        verbose_name="Montant à justifier", help_text="Écart à justifier")
    config_exercice    = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)

    class Meta:
        verbose_name = "Historique bancaire"
        unique_together = [('adherent', 'mois', 'annee')]

    @property
    def en_compte(self):
        """Total versements ou valeur saisie si disponible."""
        calcule = self.versement_tontine + self.versement_especes + self.versement_banque + self.autre_versement
        return self.en_compte_reel if self.en_compte_reel > 0 else calcule

    @property
    def montant_a_justifier(self):
        """Écart saisi ou calculé."""
        if self.montant_a_justifier_saisi != Decimal('0'):
            return self.montant_a_justifier_saisi
        return self.montant_engagement - self.en_compte


class Cheque(models.Model):
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT,
                                        related_name='cheques')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    numero          = models.CharField(max_length=50)
    montant         = models.DecimalField(max_digits=14, decimal_places=2)
    affectation     = models.CharField(max_length=200)
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)

    class Meta:
        verbose_name = "Chèque"
        ordering = ['annee', 'mois']