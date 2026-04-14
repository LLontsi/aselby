from django.db import models
from decimal import Decimal

class ContributionFoyer(models.Model):
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='contributions_foyer')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    montant_auto    = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'), help_text="150 000 FCFA si lot principal obtenu")
    don_volontaire  = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    class Meta:
        verbose_name = "Contribution foyer"
        unique_together = [('adherent', 'mois', 'annee')]
    @property
    def total(self): return self.montant_auto + self.don_volontaire
