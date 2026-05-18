from django.db import models
from decimal import Decimal
D = Decimal


class Pret(models.Model):
    EN_COURS    = 'EN_COURS'
    SOLDE       = 'SOLDE'
    LISTE_NOIRE = 'LISTE_NOIRE'
    STATUT_CHOICES = [
        (EN_COURS,    'En cours'),
        (SOLDE,       'Soldé'),
        (LISTE_NOIRE, 'Liste noire'),
    ]

    adherent          = models.ForeignKey(
        'adherents.Adherent', on_delete=models.PROTECT, related_name='prets')
    config_exercice   = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)
    montant_principal = models.DecimalField(max_digits=14, decimal_places=2)
    taux_mensuel      = models.DecimalField(max_digits=5, decimal_places=2)
    nombre_mois       = models.IntegerField()
    interet_total     = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    montant_total_du  = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'))
    date_octroi       = models.DateField()
    date_echeance     = models.DateField()
    mode_versement    = models.CharField(max_length=10,
        choices=[('BANQUE','Banque'),('ESPECES','Espèces')])
    numero_cheque     = models.CharField(max_length=50, blank=True)
    statut            = models.CharField(
        max_length=15, choices=STATUT_CHOICES, default=EN_COURS)
    nb_mois_retard    = models.IntegerField(default=0)
    montant_rembourse = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'))

    # Demande en ligne
    est_demande_membre = models.BooleanField(default=False)
    est_valide_bureau  = models.BooleanField(default=False)
    motif_demande      = models.TextField(blank=True)
    date_demande       = models.DateTimeField(null=True, blank=True)
    date_validation    = models.DateTimeField(null=True, blank=True)

    class Meta:
        verbose_name = "Prêt"
        ordering     = ['-date_octroi']

    def __str__(self):
        return (f"Prêt {self.adherent.matricule} — "
                f"{self.montant_principal:,.0f} F ({self.get_statut_display()})")

    def save(self, *args, **kwargs):
        # Recalcul auto intérêt et montant total
        self.interet_total   = (self.montant_principal
                                * (self.taux_mensuel / 100)
                                * self.nombre_mois)
        self.montant_total_du = self.montant_principal + self.interet_total
        super().save(*args, **kwargs)

    @property
    def solde_restant(self):
        return self.montant_total_du - self.montant_rembourse

    @property
    def taux_effectif(self):
        cfg = self.config_exercice
        if self.nb_mois_retard == 0:
            return self.taux_mensuel
        elif self.nb_mois_retard == 1:
            return self.taux_mensuel + cfg.majoration_retard_mois_1
        return self.taux_mensuel + cfg.majoration_retard_mois_2

    def enregistrer_remboursement(self, montant, mois, annee,
                                   mode='BANQUE', numero_cheque=''):
        """
        Enregistre un remboursement mensuel.
        Met à jour montant_rembourse et statut automatiquement.
        Lié à SaisieMonthly.remboursement_pret.
        """
        RemboursementPret.objects.update_or_create(
            pret=self, mois=mois, annee=annee,
            defaults={
                'montant':        montant,
                'mode_versement': mode,
                'numero_cheque':  numero_cheque,
            }
        )
        self.montant_rembourse = sum(
            r.montant for r in self.remboursements.all()
        ) + montant
        # Passage automatique SOLDE si entièrement remboursé
        if self.montant_rembourse >= self.montant_total_du:
            self.statut            = self.SOLDE
            self.montant_rembourse = self.montant_total_du
        self.save()


class RemboursementPret(models.Model):
    pret            = models.ForeignKey(
        Pret, on_delete=models.PROTECT, related_name='remboursements')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    montant         = models.DecimalField(max_digits=14, decimal_places=2)
    mode_versement  = models.CharField(max_length=10,
        choices=[('BANQUE','Banque'),('ESPECES','Espèces')])
    numero_cheque   = models.CharField(max_length=50, blank=True)
    penalite_retard = models.DecimalField(
        max_digits=12, decimal_places=2, default=D('0'))
    date_saisie     = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name    = "Remboursement prêt"
        unique_together = [('pret', 'mois', 'annee')]
        ordering        = ['annee', 'mois']

    def __str__(self):
        return (f"Remb. {self.pret.adherent.matricule} "
                f"{self.mois:02d}/{self.annee} — {self.montant:,.0f} F")