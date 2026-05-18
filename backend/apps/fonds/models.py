"""
apps/fonds/models.py
MouvementFonds — calculé automatiquement depuis SaisieMonthly.
Correspond à AUTREMOUVEMENT et BASECALCULINTERET.
N'est JAMAIS saisi manuellement.
"""
from django.db import models
from decimal import Decimal
import math
D = Decimal


class MouvementFonds(models.Model):
    """
    Calculé auto après chaque SaisieMonthly sauvegardée.
    Fidèle aux formules BASECALCULINTERET.xlsx col C→I.
    """
    adherent         = models.ForeignKey(
        'adherents.Adherent', on_delete=models.PROTECT,
        related_name='mouvements_fonds')
    mois             = models.IntegerField()
    annee            = models.IntegerField()
    config_exercice  = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)

    # ── Entrées ───────────────────────────────────────────────────
    capital_compose_precedent = models.DecimalField(
        max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col C = capital_compose du mois précédent")
    reconduction   = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'))
    retrait_partiel = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col E = SaisieMonthly.retrait_partiel")

    # ── Calculés et sauvegardés ───────────────────────────────────
    reste            = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col AS TABBORD = épargne brute tontine")
    epargne_nette    = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col L = reste - charges_fixes")
    fonds_roulement  = models.DecimalField(max_digits=10, decimal_places=2, default=D('0'))
    frais_exceptionnels = models.DecimalField(max_digits=10, decimal_places=2, default=D('0'))
    collation        = models.DecimalField(max_digits=10, decimal_places=2, default=D('0'))
    fonds_definitif  = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col F = capital_prec + reconduction - retrait + epargne_nette")
    base_calcul_interet = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col G = IF(fonds_def > seuil, capital_prec - retrait + epargne, 0)")
    interet_attribue = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col H = ROUNDDOWN(pool × base / total_bases, 2)")
    capital_compose  = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col I = fonds_definitif + interet_attribue")
    sanction         = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))

    date_calcul      = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name        = "Mouvement fonds"
        verbose_name_plural = "Mouvements fonds"
        unique_together     = [('adherent', 'mois', 'annee')]
        ordering            = ['annee', 'mois', 'adherent__numero_ordre']

    def __str__(self):
        return f"MVT {self.adherent.matricule} {self.mois:02d}/{self.annee}"

    # ── Méthodes de calcul (ordre exact des formules Excel) ───────
    def calculer_depuis_saisie(self, saisie):
        """
        Recalcule tout depuis une SaisieMonthly.
        Appelé automatiquement après save() de SaisieMonthly.
        """
        cfg = self.config_exercice

        # Retrait partiel depuis saisie
        self.retrait_partiel = saisie.retrait_partiel or D('0')
        self.sanction        = saisie.sanction or D('0')

        # Reste (TABBORD col AS)
        self.reste = saisie.reste  # propriété calculée dans SaisieMonthly

        # Épargne nette (col L) + charges fixes
        if self.reste <= 0:
            self.reste           = D('0')
            self.epargne_nette   = D('0')
            self.fonds_roulement = D('0')
            self.frais_exceptionnels = D('0')
            self.collation       = D('0')
        else:
            self.fonds_roulement    = cfg.fonds_roulement_mensuel
            self.frais_exceptionnels = cfg.frais_exceptionnels_mensuel
            self.collation          = cfg.collation_mensuelle
            self.epargne_nette = self.reste - cfg.charges_fixes_mensuelles

        # Fonds définitif (col F)
        self.fonds_definitif = (
            self.capital_compose_precedent
            + self.reconduction
            - self.retrait_partiel
            + self.epargne_nette
        )

        # Base calcul intérêt (col G)
        seuil = cfg.seuil_eligibilite_interets
        if self.fonds_definitif > seuil:
            self.base_calcul_interet = (
                self.capital_compose_precedent
                - self.retrait_partiel
                + self.epargne_nette
            )
        else:
            self.base_calcul_interet = D('0')

        # Capital composé (col I) — intérêt sera ajouté après répartition globale
        self.capital_compose = self.fonds_definitif + self.interet_attribue

    def appliquer_interet(self, pool_total, total_bases):
        """
        Col H = ROUNDDOWN(pool × base_i / total_bases, 2)
        Appelé depuis le service de répartition mensuelle.
        """
        if total_bases == 0 or self.base_calcul_interet == 0:
            self.interet_attribue = D('0')
        else:
            ratio = float(pool_total) / float(total_bases)
            raw   = ratio * float(self.base_calcul_interet)
            self.interet_attribue = D(str(math.floor(raw * 100) / 100))
        self.capital_compose = self.fonds_definitif + self.interet_attribue


class ReserveMensuelle(models.Model):
    """Pool d'intérêts mensuel à répartir entre adhérents éligibles."""
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    config_exercice = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)
    pool_interets             = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'))
    total_bases_eligibles     = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'))
    nb_adherents_eligibles    = models.IntegerField(default=0)
    est_reparti               = models.BooleanField(default=False)

    class Meta:
        unique_together = [('mois', 'annee', 'config_exercice')]
        ordering        = ['annee', 'mois']

    def __str__(self):
        return f"Réserve {self.mois:02d}/{self.annee} — {self.pool_interets:,.0f} F"