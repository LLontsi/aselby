from django.db import models
from decimal import Decimal


class Pret(models.Model):
    EN_COURS  = 'EN_COURS'
    SOLDE     = 'SOLDE'
    LISTE_NOIRE = 'LISTE_NOIRE'
    STATUT_CHOICES = [
        (EN_COURS, 'En cours'),
        (SOLDE, 'Soldé'),
        (LISTE_NOIRE, 'Liste noire'),
    ]

    adherent         = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='prets')
    config_exercice  = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)
    montant_principal = models.DecimalField(max_digits=14, decimal_places=2)
    taux_mensuel     = models.DecimalField(max_digits=5, decimal_places=2)  # Copié depuis config
    nombre_mois      = models.IntegerField()
    interet_total    = models.DecimalField(max_digits=12, decimal_places=2)
    montant_total_du = models.DecimalField(max_digits=14, decimal_places=2)
    date_octroi      = models.DateField()
    date_echeance    = models.DateField()
    mode_versement   = models.CharField(max_length=10, choices=[('BANQUE','Banque'),('ESPECES','Espèces')])
    numero_cheque    = models.CharField(max_length=50, blank=True)
    statut           = models.CharField(max_length=15, choices=STATUT_CHOICES, default=EN_COURS)
    nb_mois_retard   = models.IntegerField(default=0)
    montant_rembourse = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    # Demande en ligne par le membre
    est_demande_membre = models.BooleanField(default=False)
    est_valide_bureau  = models.BooleanField(default=False)
    motif_demande      = models.TextField(blank=True)
    date_demande       = models.DateTimeField(null=True, blank=True)
    date_validation    = models.DateTimeField(null=True, blank=True)

    class Meta:
        verbose_name = "Prêt"
        ordering = ['-date_octroi']

    def __str__(self):
        return f"Prêt {self.adherent.matricule} - {self.montant_principal:,.0f} FCFA ({self.get_statut_display()})"

    def save(self, *args, **kwargs):
        """Calcul automatique de l'intérêt et du montant total."""
        # Excel: interet = montant * 1% * nb_mois
        self.interet_total = self.montant_principal * (self.taux_mensuel / 100) * self.nombre_mois
        self.montant_total_du = self.montant_principal + self.interet_total
        super().save(*args, **kwargs)

    @property
    def solde_restant(self):
        return self.montant_total_du - self.montant_rembourse

    @property
    def taux_effectif(self):
        """Taux avec majoration retard selon config."""
        cfg = self.config_exercice
        if self.nb_mois_retard == 0:
            return self.taux_mensuel
        elif self.nb_mois_retard == 1:
            return self.taux_mensuel + cfg.majoration_retard_mois_1
        else:
            return self.taux_mensuel + cfg.majoration_retard_mois_2


class RemboursementPret(models.Model):
    pret            = models.ForeignKey(Pret, on_delete=models.PROTECT, related_name='remboursements')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    montant         = models.DecimalField(max_digits=14, decimal_places=2)
    mode_versement  = models.CharField(max_length=10, choices=[('BANQUE','Banque'),('ESPECES','Espèces')])
    numero_cheque   = models.CharField(max_length=50, blank=True)
    penalite_retard = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    date_saisie     = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Remboursement"
        ordering = ['annee', 'mois']
