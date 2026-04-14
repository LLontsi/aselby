from django.db import models
from decimal import Decimal


class Adherent(models.Model):
    ACTIF   = 'ACTIF'
    INACTIF = 'INACTIF'
    STATUT_CHOICES = [(ACTIF, 'Actif'), (INACTIF, 'Inactif')]

    matricule      = models.CharField(max_length=20, primary_key=True)
    numero_ordre   = models.IntegerField(unique=True)
    nom_prenom     = models.CharField(max_length=200)
    fonction       = models.CharField(max_length=200, blank=True)
    telephone1     = models.CharField(max_length=25, blank=True)
    telephone2     = models.CharField(max_length=25, blank=True)
    residence      = models.CharField(max_length=200, blank=True)
    date_adhesion  = models.DateField(null=True, blank=True)
    date_reception = models.DateField(null=True, blank=True, verbose_name="Date 1ère réception")
    statut         = models.CharField(max_length=10, choices=STATUT_CHOICES, default=ACTIF)
    photo          = models.ImageField(upload_to='adherents/', null=True, blank=True)
    poste_bureau   = models.CharField(max_length=100, blank=True, default='', verbose_name='Poste au bureau')
    notes          = models.TextField(blank=True)

    # ── CAPITAL DE DÉPART POUR LE PROCHAIN EXERCICE ──────────────────
    # Ce champ est renseigné en fin d'exercice par le bureau.
    # Il contient le capital composé final (NOUVEAU FONDS de la fiche cassation).
    # Pour 2025 : valeurs importées depuis MVTDEC244 (Excel).
    # Pour 2026 et suivants : rempli automatiquement à la clôture de l'exercice.
    capital_depart_exercice = models.DecimalField(
        max_digits=14, decimal_places=2,
        default=Decimal('0'),
        verbose_name="Capital de départ (prochain exercice)",
        help_text=(
            "Capital composé fin d'exercice précédent. "
            "Sert de base au calcul du fonds définitif mois 1 de l'exercice suivant. "
            "Mis à jour automatiquement à la clôture."
        )
    )

    date_creation     = models.DateTimeField(auto_now_add=True)
    date_modification = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Adhérent"
        verbose_name_plural = "Adhérents"
        ordering = ['numero_ordre']

    def __str__(self):
        return f"{self.matricule} - {self.nom_prenom}"

    @property
    def est_actif(self):
        return self.statut == self.ACTIF

    def get_parts_tontine(self, annee):
        """Retourne les parts de cet adhérent pour une année donnée."""
        return self.participations_tontine.filter(
            session__annee=annee
        ).select_related('session__niveau')
