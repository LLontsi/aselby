from django.db import models
from django.core.validators import MinValueValidator
from decimal import Decimal


class ConfigExercice(models.Model):
    """
    Configuration complète d'un exercice annuel.
    L'admin fixe TOUTES ces valeurs en début d'exercice avant ouverture.
    Une fois ouvert, les paramètres sont gelés.
    """
    annee       = models.IntegerField(unique=True, verbose_name="Année")
    est_ouvert  = models.BooleanField(default=False, verbose_name="Exercice ouvert")
    date_ouverture = models.DateField(null=True, blank=True)
    date_cloture   = models.DateField(null=True, blank=True)

    # ── TONTINES ──────────────────────────────────────────────────────
    taux_t35 = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('35000'),
        verbose_name="Taux mensuel T35 (FCFA)")
    versement_t35 = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('60000'),
        verbose_name="Versement mensuel par part T35 (FCFA)")
    taux_t75 = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('75000'),
        verbose_name="Taux mensuel T75 (FCFA)")
    taux_t100 = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('100000'),
        verbose_name="Taux mensuel T100 (FCFA)")
    diviseur_interet_t35 = models.IntegerField(
        default=35, verbose_name="Diviseur intérêts T35 (nb parts)")

    # ── FONDS DE CAISSE — cotisations mensuelles fixées par l'admin ──
    seuil_eligibilite_interets = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('900000'),
        verbose_name="Seuil éligibilité intérêts fonds (FCFA)")
    fonds_roulement_mensuel = models.DecimalField(
        max_digits=10, decimal_places=2, default=Decimal('1000'),
        verbose_name="Fonds de roulement / adhérent / mois (FCFA)")
    frais_exceptionnels_mensuel = models.DecimalField(
        max_digits=10, decimal_places=2, default=Decimal('1000'),
        verbose_name="Frais exceptionnels / adhérent / mois (FCFA)")
    collation_mensuelle = models.DecimalField(
        max_digits=10, decimal_places=2, default=Decimal('4000'),
        verbose_name="Collation / adhérent / mois (FCFA)")
    mutuelle_mensuelle = models.DecimalField(
        max_digits=10, decimal_places=2, default=Decimal('0'),
        verbose_name="Cotisation mutuelle mensuelle / adhérent (FCFA)")
    montant_inscription = models.DecimalField(
        max_digits=10, decimal_places=2, default=Decimal('25000'),
        verbose_name="Frais d'inscription nouvel adhérent (FCFA)")
    complement_fonds_fin_exercice = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('100000'),
        verbose_name="Complément fonds fin d'exercice (FCFA)")

    # ── PÉNALITÉS ────────────────────────────────────────────────────
    penalite_especes = models.DecimalField(
        max_digits=10, decimal_places=2, default=Decimal('3000'),
        verbose_name="Pénalité versement espèces (FCFA)")
    penalite_especes_active = models.BooleanField(
        default=True, verbose_name="Pénalité espèces active")
    pourcentage_penalite_echec = models.DecimalField(
        max_digits=5, decimal_places=2, default=Decimal('20'),
        verbose_name="% pénalité échec tontine")
    nb_echecs_max_avant_liste_rouge = models.IntegerField(
        default=3, verbose_name="Nb max d'échecs avant liste rouge")

    # ── PRÊTS ────────────────────────────────────────────────────────
    taux_interet_pret_mensuel = models.DecimalField(
        max_digits=5, decimal_places=2, default=Decimal('1'),
        verbose_name="Taux intérêt prêt (%/mois)")
    majoration_retard_mois_1 = models.DecimalField(
        max_digits=5, decimal_places=2, default=Decimal('5'),
        verbose_name="Majoration retard prêt mois 1 (%)")
    majoration_retard_mois_2 = models.DecimalField(
        max_digits=5, decimal_places=2, default=Decimal('10'),
        verbose_name="Majoration retard prêt mois 2 (%)")
    montant_min_pret = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('0'),
        verbose_name="Montant minimum prêt (FCFA)")
    montant_max_pret = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('2000000'),
        verbose_name="Montant maximum prêt (FCFA)")

    # ── FOYER ────────────────────────────────────────────────────────
    contribution_foyer_lot_principal = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('150000'),
        verbose_name="Contribution foyer sur lot principal (FCFA)")

    # ── MUTUELLE ─────────────────────────────────────────────────────
    complement_mutuelle_fin_exercice = models.DecimalField(
        max_digits=12, decimal_places=2, default=Decimal('30000'),
        verbose_name="Complément mutuelle fin d'exercice (FCFA)")

    class Meta:
        verbose_name = "Configuration exercice"
        verbose_name_plural = "Configurations exercices"
        ordering = ['-annee']

    def __str__(self):
        return f"Exercice {self.annee} ({'Ouvert' if self.est_ouvert else 'Fermé'})"

    @property
    def charges_fixes_mensuelles(self):
        """Total des charges déduites du reste mensuel (fonds + frais + collation)."""
        return (
            self.fonds_roulement_mensuel
            + self.frais_exceptionnels_mensuel
            + self.collation_mensuelle
        )

    @classmethod
    def get_exercice_courant(cls):
        # Toujours retourner le dernier exercice ouvert (annee la plus récente)
        return cls.objects.filter(est_ouvert=True).order_by('-annee').first()


# ══════════════════════════════════════════════════════════════════════
# DÉPENSES PAR RUBRIQUE — sorties avec libellés (gérées par l'admin)
# ══════════════════════════════════════════════════════════════════════

class DepenseFondsRoulement(models.Model):
    """Sortie du fonds de roulement. Ex: TAXI BANQUE 6 000 F, COMMISSION 30 000 F"""
    config_exercice = models.ForeignKey(
        ConfigExercice, on_delete=models.PROTECT,
        related_name='depenses_fonds_roulement')
    date    = models.DateField(verbose_name="Date")
    libelle = models.CharField(max_length=300, verbose_name="Libellé")
    montant = models.DecimalField(max_digits=14, decimal_places=2,
        validators=[MinValueValidator(Decimal('0'))], verbose_name="Montant (FCFA)")
    date_saisie = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Dépense fonds de roulement"
        ordering = ['-date']

    def __str__(self):
        return f"{self.date} — {self.libelle} ({self.montant:,.0f} F)"


class DepenseFraisExceptionnels(models.Model):
    """Sortie des frais exceptionnels. Ex: GRATIFICATION 2024 = 200 000 F"""
    config_exercice = models.ForeignKey(
        ConfigExercice, on_delete=models.PROTECT,
        related_name='depenses_frais_exceptionnels')
    date    = models.DateField(verbose_name="Date")
    libelle = models.CharField(max_length=300, verbose_name="Libellé")
    montant = models.DecimalField(max_digits=14, decimal_places=2,
        validators=[MinValueValidator(Decimal('0'))], verbose_name="Montant (FCFA)")
    date_saisie = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Dépense frais exceptionnels"
        ordering = ['-date']

    def __str__(self):
        return f"{self.date} — {self.libelle} ({self.montant:,.0f} F)"


class DepenseCollation(models.Model):
    """Sortie de la collation. Ex: Collation séance = 100 000 F"""
    config_exercice = models.ForeignKey(
        ConfigExercice, on_delete=models.PROTECT,
        related_name='depenses_collation')
    date    = models.DateField(verbose_name="Date de la séance")
    libelle = models.CharField(max_length=300, default="COLLATION SEANCE", verbose_name="Libellé")
    montant = models.DecimalField(max_digits=14, decimal_places=2,
        validators=[MinValueValidator(Decimal('0'))], verbose_name="Montant (FCFA)")
    date_saisie = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Dépense collation"
        ordering = ['-date']

    def __str__(self):
        return f"{self.date} — {self.libelle} ({self.montant:,.0f} F)"


class DepenseFoyer(models.Model):
    """Virement foyer vers compte conseil. Ex: VERSE DANS LE COMPTE CONSEIL 5 000 000 F"""
    config_exercice = models.ForeignKey(
        ConfigExercice, on_delete=models.PROTECT,
        related_name='depenses_foyer')
    date    = models.DateField(verbose_name="Date du virement")
    libelle = models.CharField(max_length=300, default="VERSE DANS LE COMPTE CONSEIL",
                                verbose_name="Libellé")
    montant = models.DecimalField(max_digits=14, decimal_places=2,
        validators=[MinValueValidator(Decimal('0'))], verbose_name="Montant (FCFA)")
    date_saisie = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Dépense foyer"
        ordering = ['-date']

    def __str__(self):
        return f"{self.date} — {self.libelle} ({self.montant:,.0f} F)"


class AgioBancaire(models.Model):
    """Agios et intérêts créditeurs mensuels — saisis depuis le relevé bancaire."""
    config_exercice  = models.ForeignKey(
        ConfigExercice, on_delete=models.PROTECT, related_name='agios')
    mois  = models.IntegerField(verbose_name="Mois (1-12)")
    annee = models.IntegerField(verbose_name="Année")
    montant_agio       = models.DecimalField(max_digits=10, decimal_places=2,
        default=Decimal('0'), verbose_name="Agio / frais bancaires (FCFA)")
    interet_crediteur  = models.DecimalField(max_digits=10, decimal_places=2,
        default=Decimal('0'), verbose_name="Intérêt créditeur banque (FCFA)")
    date_saisie = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Agio bancaire"
        unique_together = [('mois', 'annee')]
        ordering = ['annee', 'mois']

    def __str__(self):
        return f"Agios {self.mois:02d}/{self.annee}"


class AutreDepense(models.Model):
    """Dépenses libres hors rubriques. Ex: SOLDE TERRAIN 350 000 F, NOTAIRE 50 000 F"""
    config_exercice = models.ForeignKey(
        ConfigExercice, on_delete=models.PROTECT, related_name='autres_depenses')
    date    = models.DateField(verbose_name="Date")
    libelle = models.CharField(max_length=300, verbose_name="Libellé")
    montant = models.DecimalField(max_digits=14, decimal_places=2,
        validators=[MinValueValidator(Decimal('0'))], verbose_name="Montant (FCFA)")
    date_saisie = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Autre dépense"
        ordering = ['-date']

    def __str__(self):
        return f"{self.date} — {self.libelle} ({self.montant:,.0f} F)"