from django.db import models
from decimal import Decimal


class ListeRouge(models.Model):
    """Adhérent en liste rouge - totalité des avoirs en garantie."""
    adherent        = models.OneToOneField('adherents.Adherent', on_delete=models.PROTECT, related_name='liste_rouge')
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    date_entree     = models.DateField(auto_now_add=True)
    motif           = models.TextField()
    montant_dette   = models.DecimalField(max_digits=14, decimal_places=2)
    montant_garantie = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    est_solde       = models.BooleanField(default=False)
    date_solde      = models.DateField(null=True, blank=True)

    class Meta:
        verbose_name = "Liste rouge"

    def __str__(self):
        statut = "Soldé" if self.est_solde else "En cours"
        return f"{self.adherent.matricule} - {self.montant_dette:,.0f} FCFA ({statut})"


class RemboursementDette(models.Model):
    liste_rouge   = models.ForeignKey(ListeRouge, on_delete=models.PROTECT, related_name='remboursements')
    date          = models.DateField()
    montant       = models.DecimalField(max_digits=14, decimal_places=2)
    observations  = models.TextField(blank=True)

    class Meta:
        ordering = ['date']

    def __str__(self):
        return f"Remboursement {self.liste_rouge.adherent.matricule} - {self.montant:,.0f} FCFA"
