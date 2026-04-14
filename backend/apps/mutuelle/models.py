from django.db import models
from decimal import Decimal

class CotisationMutuelle(models.Model):
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='cotisations_mutuelle')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    montant         = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    class Meta:
        verbose_name = "Cotisation mutuelle"
        unique_together = [('adherent', 'mois', 'annee')]
        ordering = ['annee', 'mois']

class AideMutuelle(models.Model):
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='aides_mutuelle')
    date            = models.DateField()
    evenement       = models.CharField(max_length=300)
    montant         = models.DecimalField(max_digits=12, decimal_places=2)
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    class Meta:
        verbose_name = "Aide mutuelle"
        ordering = ['-date']
    def __str__(self): return f"{self.adherent.matricule} — {self.evenement} ({self.date})"
