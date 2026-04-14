from django.db import models
from django.core.validators import MinValueValidator
from decimal import Decimal


class NiveauTontine(models.Model):
    """T35, T75, T100 - configuré par exercice."""
    config_exercice = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT,
        related_name='niveaux_tontine'
    )
    code         = models.CharField(max_length=10)  # T35, T75, T100
    taux_mensuel = models.DecimalField(max_digits=12, decimal_places=2)
    # Pour T35 uniquement : versement != taux
    versement_mensuel_par_part = models.DecimalField(max_digits=12, decimal_places=2)
    diviseur_interet = models.IntegerField(default=1)

    class Meta:
        verbose_name = "Niveau tontine"
        unique_together = [('config_exercice', 'code')]
        ordering = ['taux_mensuel']

    def __str__(self):
        return f"{self.code} - {self.taux_mensuel:,.0f} FCFA/mois"

    @property
    def lot_annuel(self):
        return self.taux_mensuel * 12

    @property
    def a_fonds_complementaire(self):
        """Seule la T35 a un fonds complémentaire intégré."""
        return self.versement_mensuel_par_part > self.taux_mensuel


class SessionTontine(models.Model):
    """Une session = un mois pour un niveau donné."""
    niveau  = models.ForeignKey(NiveauTontine, on_delete=models.PROTECT, related_name='sessions')
    mois    = models.IntegerField()
    annee   = models.IntegerField()
    date_seance = models.DateField()
    # Intérêt décidé par le bureau pour cette session
    montant_interet_bureau = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('0'),
        verbose_name="Montant intérêts décidé par le bureau"
    )
    est_cloturee = models.BooleanField(default=False)

    class Meta:
        verbose_name = "Session tontine"
        unique_together = [('niveau', 'mois', 'annee')]
        ordering = ['annee', 'mois']

    def __str__(self):
        return f"{self.niveau.code} - {self.mois:02d}/{self.annee}"

    @property
    def interet_par_adherent(self):
        """ROUNDDOWN(montant_bureau / diviseur, 0)."""
        if self.montant_interet_bureau and self.niveau.diviseur_interet:
            import math
            return math.floor(
                float(self.montant_interet_bureau) / self.niveau.diviseur_interet
            )
        return 0


class ParticipationTontine(models.Model):
    """Participation mensuelle d'un adhérent à une session."""
    BANQUE  = 'BANQUE'
    ESPECES = 'ESPECES'
    ECHEC   = 'ECHEC'
    MODE_CHOICES = [(BANQUE, 'Banque'), (ESPECES, 'Espèces'), (ECHEC, 'Échec')]

    session         = models.ForeignKey(SessionTontine, on_delete=models.PROTECT, related_name='participations')
    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT, related_name='participations_tontine')
    nombre_parts    = models.IntegerField(default=1)
    mode_versement  = models.CharField(max_length=10, choices=MODE_CHOICES, default=ECHEC)
    montant_verse   = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    penalite_especes = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    penalite_echec  = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    nb_echecs_cumules = models.IntegerField(default=0)
    en_liste_rouge  = models.BooleanField(default=False)
    complement_epargne = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    numero_cheque   = models.CharField(max_length=50, blank=True)
    # Lots
    a_obtenu_lot_principal = models.BooleanField(default=False)
    montant_lot_principal  = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    interet_lot_principal  = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    # Petit lot
    vente_petit_lot  = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    interet_petit_lot = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    remboursement_petit_lot = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    mode_remboursement_petit_lot = models.CharField(max_length=10, choices=MODE_CHOICES, blank=True)

    class Meta:
        verbose_name = "Participation tontine"
        unique_together = [('session', 'adherent')]

    def __str__(self):
        return f"{self.adherent.matricule} - {self.session} ({self.nombre_parts} part(s))"

    @property
    def montant_tontine_du(self):
        """Montant de tontine pur dû selon les parts et le mode."""
        if self.mode_versement == self.ECHEC:
            return Decimal('0')
        return self.session.niveau.taux_mensuel * self.nombre_parts

    @property
    def montant_engagement_total(self):
        """Total des engagements mensuels de cet adhérent."""
        return (
            self.montant_tontine_du * self.session.niveau.nombre_parts_versement
            + self.complement_epargne
            + self.penalite_especes
            + self.remboursement_petit_lot
        )

    @property
    def reste(self):
        """Reste = versement_total - taux_pur. Base du fonds complémentaire."""
        if self.mode_versement == self.ECHEC:
            return Decimal('0')
        taux_pur = self.session.niveau.taux_mensuel * self.nombre_parts
        return self.montant_verse + self.complement_epargne - taux_pur
