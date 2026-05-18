"""
apps/banque/models.py
Relevé bancaire mensuel — TABBHISTOBQUE
Saisi EN PREMIER, avant la saisie mensuelle.
Colonnes manuelles : versement_banque (E), versement_especes (D),
                     autre_versement (F), agio (H)
Colonnes auto      : montant_engagement (G) = D+E+F
"""
from django.db import models
from decimal import Decimal
D = Decimal


class ReleveBancaire(models.Model):
    """
    TABBHISTOBQUE — une ligne par adhérent par mois.
    Saisie manuelle : versements réels du relevé bancaire + agio.
    """
    adherent        = models.ForeignKey(
        'adherents.Adherent', on_delete=models.PROTECT,
        related_name='releves_bancaires')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    config_exercice = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)

    # ── Champs manuels ────────────────────────────────────────────
    versement_banque  = models.DecimalField(
        max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col E — chèques/virements")
    versement_especes = models.DecimalField(
        max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col D — espèces")
    autre_versement   = models.DecimalField(
        max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col F — versements exceptionnels")
    agio              = models.DecimalField(
        max_digits=10, decimal_places=2, default=D('0'),
        help_text="Col H — frais bancaires du mois")

    # Soumission membre depuis espace personnel
    est_valide_membre  = models.BooleanField(default=False,
        help_text="Soumis par le membre depuis son espace")
    est_valide_bureau  = models.BooleanField(default=True,
        help_text="Validé par le bureau (False = en attente)")
    date_saisie       = models.DateTimeField(auto_now_add=True)
    date_modification = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name        = "Relevé bancaire"
        verbose_name_plural = "Relevés bancaires"
        unique_together     = [('adherent', 'mois', 'annee')]
        ordering            = ['annee', 'mois', 'adherent__numero_ordre']

    def __str__(self):
        return f"Relevé {self.adherent.matricule} {self.mois:02d}/{self.annee}"

    @property
    def montant_engagement(self):
        """Col G = D + E + F"""
        return self.versement_banque + self.versement_especes + self.autre_versement

    @property
    def mode_versement(self):
        """BANQUE si banque > 0, ESPECES si espèces > 0, sinon ECHEC"""
        if self.versement_banque > 0:
            return 'BANQUE'
        if self.versement_especes > 0:
            return 'ESPECES'
        return 'ECHEC'


class AgioBancaire(models.Model):
    """Agio global de l'association (pas par adhérent)."""
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    config_exercice = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)
    numero          = models.CharField(max_length=50, blank=True)
    montant         = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    affectation     = models.CharField(max_length=200, blank=True)

    class Meta:
        ordering = ['annee', 'mois']

    def __str__(self):
        return f"Agio {self.mois:02d}/{self.annee} — {self.montant:,.0f} F"