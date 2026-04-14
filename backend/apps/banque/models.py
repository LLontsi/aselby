from django.db import models
from decimal import Decimal

class HistoriqueBancaire(models.Model):
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='historique_bancaire')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    versement_tontine = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    versement_especes = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    versement_banque  = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    autre_versement   = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    montant_engagement = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    agio             = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    config_exercice  = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    class Meta:
        verbose_name = "Historique bancaire"
        unique_together = [('adherent', 'mois', 'annee')]
    @property
    def en_compte(self): return self.versement_tontine + self.versement_especes + self.versement_banque + self.autre_versement
    @property
    def montant_a_justifier(self): return self.montant_engagement - self.en_compte

class Cheque(models.Model):
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='cheques')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    numero          = models.CharField(max_length=50)
    montant         = models.DecimalField(max_digits=14, decimal_places=2)
    affectation     = models.CharField(max_length=200)
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    class Meta:
        verbose_name = "Chèque"
        ordering = ['annee', 'mois']
