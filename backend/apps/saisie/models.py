"""
apps/saisie/models.py
SaisieMonthly — TABBORD mensuel.
Saisi APRÈS ReleveBancaire (TABBHISTOBQUE).

Champs manuels : 28 colonnes TABBORD sans formule
Champs auto    : 34 colonnes calculées → exposées comme propriétés

Formules fidèles au TABBORD Excel 2025.
"""
from django.db import models
from decimal import Decimal
import math
D = Decimal

BANQUE  = 'BANQUE'
ESPECES = 'ESPECES'
ECHEC   = 'ECHEC'
MODE_CHOICES = [(BANQUE,'Banque'),(ESPECES,'Espèces'),(ECHEC,'Échec')]


class SaisieMonthly(models.Model):
    """
    Une ligne par adhérent par mois.
    Source de vérité pour TABBORD, AUTREMOUVEMENT, TONTINE exports.
    """
    adherent        = models.ForeignKey(
        'adherents.Adherent', on_delete=models.PROTECT,
        related_name='saisies_monthly')
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    config_exercice = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)

    # ════════════════════════════════════════════════════════════
    # CHAMPS MANUELS — TONTINES (obligatoires chaque mois)
    # ════════════════════════════════════════════════════════════
    nbre_lots_t60  = models.IntegerField(default=0, help_text="TABBORD col J — nbre parts T60")
    nbre_lots_t75  = models.IntegerField(default=0, help_text="TABBORD col L — nbre parts T75")
    nbre_lots_t100 = models.IntegerField(default=0, help_text="TABBORD col N — nbre parts T100")

    # Lots obtenus (achat du lot = intérêt offert au bureau)
    achat_lot_t60   = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col AC — intérêt offert sur lot T60")
    achat_lot_t75   = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col AD — intérêt offert sur lot T75")
    achat_lot_t100  = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col AE — intérêt offert sur lot T100")
    mode_paiement_lot = models.CharField(max_length=10, choices=MODE_CHOICES,
        blank=True, help_text="Col AF")

    # Vente petit lot (adhérent revend son droit avant l'échéance)
    vente_petit_lot_t60  = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col AI")
    vente_petit_lot_t75  = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col AJ")
    vente_petit_lot_t100 = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col AK")
    interet_petit_lot_t60  = models.DecimalField(max_digits=10, decimal_places=2, default=D('0'),
        help_text="Col AL")
    interet_petit_lot_t75  = models.DecimalField(max_digits=10, decimal_places=2, default=D('0'),
        help_text="Col AM")
    interet_petit_lot_t100 = models.DecimalField(max_digits=10, decimal_places=2, default=D('0'),
        help_text="Col AN")
    mode_remb_petit_lot = models.CharField(max_length=10, choices=MODE_CHOICES,
        blank=True, help_text="Col AP")

    # ════════════════════════════════════════════════════════════
    # CHAMPS MANUELS — PRÊTS
    # ════════════════════════════════════════════════════════════
    remboursement_pret = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col S — remboursement prêt ce mois")
    mode_remb_pret     = models.CharField(max_length=10, choices=MODE_CHOICES,
        blank=True, help_text="Col T")
    pret_fonds         = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col Z — nouveau prêt octroyé ce mois")
    mode_paiement_pret = models.CharField(max_length=10, choices=MODE_CHOICES,
        blank=True, help_text="Col AA")
    nbre_mois_pret     = models.IntegerField(default=0,
        help_text="Col BG — nombre de mois écoulés (pour calcul intérêt prêt)")

    # ════════════════════════════════════════════════════════════
    # CHAMPS MANUELS — VERSEMENTS SPÉCIAUX
    # ════════════════════════════════════════════════════════════
    complement_epargne = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col G — peut être négatif (malus)")
    montant_especes    = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col BA — montant espèces cas particuliers")
    numero_cheque      = models.CharField(max_length=50, blank=True,
        help_text="Col BE — numéro chèque effectif ou 'OK'")

    # ════════════════════════════════════════════════════════════
    # CHAMPS MANUELS — DÉPENSES
    # ════════════════════════════════════════════════════════════
    libelle_depense  = models.CharField(max_length=200, blank=True,
        help_text="Col BO — ex: ACHAT FOURNITURE")
    compte_depense   = models.CharField(max_length=100, blank=True,
        help_text="Col BP — ex: FONDS DE ROULEMENT")
    montant_depense  = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col BQ")
    depense_collation = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col BK — dépense collation (rare)")

    # ════════════════════════════════════════════════════════════
    # CHAMPS MANUELS — CAS EXCEPTIONNELS
    # ════════════════════════════════════════════════════════════
    sanction            = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col V")
    inscription         = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col W — nouvel adhérent")
    retrait_partiel     = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Col AV")
    mutuelle            = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col AT")
    contribution_foyer  = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col BB — contribution foyer volontaire")
    montant_lot_t75  = models.DecimalField(
        max_digits=12, decimal_places=2, default=0,
        verbose_name="Montant lot T75 reçu",
        help_text="Col AS TABBORD: T75.montant_lot_principal reçu ce mois"
    )
    montant_lot_t100 = models.DecimalField(
        max_digits=12, decimal_places=2, default=0,
        verbose_name="Montant lot T100 reçu",
        help_text="Col AT TABBORD: T100.montant_lot_principal reçu ce mois"
    )
    # Bonus malus saisi manuellement
    bonus_malus_saisi = models.DecimalField(
        max_digits=12, decimal_places=2, default=0,
        verbose_name="Bonus Malus (saisi manuellement)",
        help_text="Positif=excédent, négatif=déficit/prêt. Si 0: calculé auto."
    )
    # Champs de saisie manuelle (priment sur le calcul automatique)
    penalite_versement_especes_saisi = models.DecimalField(
        max_digits=12, decimal_places=2, default=0,
        verbose_name="Pénalité versement espèces (saisie)",
        help_text="Si 0: calculé automatiquement depuis la config. Sinon: valeur saisie."
    )
    interet_pret_saisi = models.DecimalField(
        max_digits=12, decimal_places=2, default=0,
        verbose_name="Intérêt prêt fonds (saisi)",
        help_text="Si 0: calculé automatiquement (taux × remb × nb_mois). Sinon: valeur saisie."
    )
    penalite_pret_fonds = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Col BL")

    date_saisie       = models.DateTimeField(auto_now_add=True)
    date_modification = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name        = "Saisie mensuelle"
        verbose_name_plural = "Saisies mensuelles"
        unique_together     = [('adherent', 'mois', 'annee')]
        ordering            = ['annee', 'mois', 'adherent__numero_ordre']

    def __str__(self):
        return f"Saisie {self.adherent.matricule} {self.mois:02d}/{self.annee}"

    # ════════════════════════════════════════════════════════════
    # PROPRIÉTÉS CALCULÉES AUTO — fidèles aux formules TABBORD
    # ════════════════════════════════════════════════════════════

    def _releve(self):
        """Relevé bancaire du même mois (peut être None)."""
        from apps.banque.models import ReleveBancaire
        try:
            return ReleveBancaire.objects.get(
                adherent=self.adherent, mois=self.mois, annee=self.annee)
        except ReleveBancaire.DoesNotExist:
            return None

    @property
    def versement_banque(self):
        """Col D — depuis ReleveBancaire"""
        r = self._releve()
        return r.versement_banque if r else D('0')

    @property
    def versement_especes(self):
        """Col E — depuis ReleveBancaire"""
        r = self._releve()
        return r.versement_especes if r else D('0')

    @property
    def autre_versement(self):
        """Col F — depuis ReleveBancaire"""
        r = self._releve()
        return r.autre_versement if r else D('0')

    @property
    def mode_paiement_tontine(self):
        """
        Col I = IF(banque<>0,'BANQUE', IF(especes<>0,'ESPECES','ECHEC'))
        Sauf override manuel (mode_remb_pret utilisé comme proxy si ECHEC forcé)
        """
        if self.versement_banque > 0:
            return BANQUE
        if self.versement_especes > 0:
            return ESPECES
        return ECHEC

    @property
    def penalite_versement_especes(self):
        """Col H = valeur saisie si != 0, sinon IF(mode='ESPECES', penalite_config, 0)"""
        if self.penalite_versement_especes_saisi and self.penalite_versement_especes_saisi > 0:
            return D(str(self.penalite_versement_especes_saisi))
        if self.mode_paiement_tontine == ESPECES:
            return self.config_exercice.penalite_especes
        return D('0')

    @property
    def tontine_t60(self):
        """Col K = IF(mode!='ECHEC', nbre_lots_t60 × versement_t60, 0)"""
        if self.mode_paiement_tontine == ECHEC:
            return D('0')
        return D(str(self.nbre_lots_t60)) * self.config_exercice.versement_t35

    @property
    def tontine_t75(self):
        """Col M"""
        if self.mode_paiement_tontine == ECHEC:
            return D('0')
        return D(str(self.nbre_lots_t75)) * self.config_exercice.taux_t75

    @property
    def tontine_t100(self):
        """Col O"""
        if self.mode_paiement_tontine == ECHEC:
            return D('0')
        return D(str(self.nbre_lots_t100)) * self.config_exercice.taux_t100

    @property
    def tontine_mois(self):
        """Col Y = K + M + O"""
        return self.tontine_t60 + self.tontine_t75 + self.tontine_t100

    @property
    def remboursement_petit_lot_t60(self):
        """Col P — calculé depuis mois précédents (montant lot + intérêt)"""
        # En 2025 = TONTFEV.H + TONTFEV.I (mois précédent)
        # Simplification : géré dans le service de calcul, pas ici
        return D('0')

    @property
    def remboursement_petit_lot_t75(self):
        return D('0')

    @property
    def remboursement_petit_lot_t100(self):
        return D('0')

    @property
    def montant_engagement(self):
        """
        Col U = G + H + S + R + P + O + M + K + Q + V + W
        = complement_epargne + penalite_especes + remb_pret
          + remb_petit_lots + tontines + sanction + inscription
        """
        return (
            self.complement_epargne
            + self.penalite_versement_especes
            + self.remboursement_pret
            + self.remboursement_petit_lot_t60
            + self.remboursement_petit_lot_t75
            + self.remboursement_petit_lot_t100
            + self.tontine_t60
            + self.tontine_t75
            + self.tontine_t100
            + self.sanction
            + self.inscription
        )

    @property
    def bonus_malus(self):
        """Col C TABBORD = saisi si != 0, sinon calcul: (D+E+F) - X - G - I - S"""
        if self.bonus_malus_saisi and self.bonus_malus_saisi != 0:
            return D(str(self.bonus_malus_saisi))
        # Calcul automatique
        from apps.banque.models import ReleveBancaire
        try:
            r = ReleveBancaire.objects.get(
                adherent=self.adherent, mois=self.mois,
                annee=self.annee, config_exercice=self.config_exercice)
            D_bq  = D(str(r.versement_banque  or 0))
            E_esp = D(str(r.versement_especes or 0))
            F_aut = D(str(r.autre_versement   or 0))
        except Exception:
            D_bq = E_esp = F_aut = D('0')
        X_ton = self.tontine_mois
        G_comp= D(str(self.complement_epargne or 0))
        I_pen = self.penalite_versement_especes
        S_pret= D(str(self.remboursement_pret  or 0))
        return (D_bq + E_esp + F_aut) - X_ton - G_comp - I_pen - S_pret
    @property
    def penalite_retard_tontine(self):
        """Col AQ = IF(mode='ECHEC', tontine_mois × 20%, 0)"""
        if self.mode_paiement_tontine == ECHEC:
            cfg = self.config_exercice
            return self.tontine_mois * cfg.pourcentage_penalite_echec / 100
        return D('0')

    @property
    def penalite_echec_tontine(self):
        """Col AR = même calcul que AQ"""
        return self.penalite_retard_tontine

    @property
    def reste(self):
        """
        Col AS = IF(mode='ECHEC', 0, max(0, tontine_t60 + complement_epargne - taux_pur_t60))
        Si négatif (ex. ASELBY sans T60), retourner 0.
        """
        if self.mode_paiement_tontine == ECHEC:
            return D('0')
        taux_pur = self.config_exercice.taux_t35
        val = self.tontine_t60 + self.complement_epargne - taux_pur
        return val if val > D('0') else D('0')

    @property
    def contribution_foyer_auto(self):
        """
        Col BB = IF(achat_lot_principal > 0, montant_config, 0)
        Contribution foyer automatique si lot principal obtenu
        """
        a_lot = (self.achat_lot_t60 > 0
                 or self.achat_lot_t75 > 0
                 or self.achat_lot_t100 > 0)
        if a_lot:
            return self.config_exercice.contribution_foyer_lot_principal
        return D('0')

    @property
    def montant_t60(self):
        """
        Col AW = IF(achat_lot=0, vente_petit_lot,
                    (1×taux×12) + vente_petit_lot - achat_lot)
        """
        if self.achat_lot_t60 == 0:
            return self.vente_petit_lot_t60
        taux_annuel = D(str(1)) * self.config_exercice.taux_t35 * 12
        return taux_annuel + self.vente_petit_lot_t60 - self.achat_lot_t60

    @property
    def montant_t75(self):
        if self.achat_lot_t75 == 0:
            return self.vente_petit_lot_t75
        taux_annuel = D(str(1)) * self.config_exercice.taux_t75 * 12
        return taux_annuel + self.vente_petit_lot_t75 - self.achat_lot_t75

    @property
    def montant_t100(self):
        if self.achat_lot_t100 == 0:
            return self.vente_petit_lot_t100
        taux_annuel = D(str(1)) * self.config_exercice.taux_t100 * 12
        return taux_annuel + self.vente_petit_lot_t100 - self.achat_lot_t100

    @property
    def interet_pret(self):
        """Intérêt sur prêt fonds — valeur saisie si != 0, sinon calcul auto"""
        if self.interet_pret_saisi and self.interet_pret_saisi > 0:
            return D(str(self.interet_pret_saisi))
        if self.pret_fonds == 0 or not self.nbre_mois_pret:
            return D('0')
        taux = self.config_exercice.taux_interet_pret_mensuel or D('0.01')
        return (D(str(self.pret_fonds)) * D(str(taux)) * D(str(self.nbre_mois_pret))).quantize(D('0.01'))
    @property
    def montant_cheque(self):
        """
        Col AZ = montant_t60 + montant_t75 + montant_t100
                 - interet_pret + pret_fonds
                 + depenses diverses
        """
        return (
            self.montant_t60
            + self.montant_t75
            + self.montant_t100
            - self.interet_pret
            + self.pret_fonds
            + self.montant_depense
            + self.depense_collation
        )

    @property
    def montant_cheque_effectif(self):
        """
        Col BD = IF(montant_cheque=0,
                    montant_cheque - montant_especes - autres_mode,
                    montant_cheque - montant_especes - contribution_foyer - autres_mode)
        """
        if self.montant_cheque == 0:
            return self.montant_cheque - self.montant_especes
        return (self.montant_cheque
                - self.montant_especes
                - self.contribution_foyer_auto)

    # Dépenses auto depuis config (parts fixes mensuelles)
    @property
    def depense_fonds_roulement(self):
        if self.reste > 0:
            return self.config_exercice.fonds_roulement_mensuel
        return D('0')

    @property
    def depense_frais_exceptionnel(self):
        if self.reste > 0:
            return self.config_exercice.frais_exceptionnels_mensuel
        return D('0')

    @property
    def depense_fonds_mutuel(self):
        if self.reste > 0:
            return self.config_exercice.collation_mensuelle
        return D('0')

    @property
    def depenses_auto_total(self):
        return (self.depense_fonds_roulement
                + self.depense_frais_exceptionnel
                + self.depense_fonds_mutuel)