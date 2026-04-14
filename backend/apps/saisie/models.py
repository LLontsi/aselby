from django.db import models
from decimal import Decimal


class TableauDeBord(models.Model):
    """
    Source unique de saisie mensuelle.
    Chaque ligne = un adhérent pour un mois donné.

    RÈGLE PÉNALITÉ ESPÈCES :
    - Ce n'est PAS automatique : c'est l'admin qui DÉCIDE d'appliquer ou non
    - Le champ `penalite_especes_appliquee` est coché par l'admin dans le formulaire
    - La valeur de la pénalité vient de config.penalite_especes (3 000 F) si cochée
    """
    BANQUE  = 'BANQUE'
    ESPECES = 'ESPECES'
    ECHEC   = 'ECHEC'
    MODE_CHOICES = [(BANQUE,'Banque'),(ESPECES,'Espèces'),(ECHEC,'Échec')]

    adherent        = models.ForeignKey('adherents.Adherent', on_delete=models.PROTECT)
    mois            = models.IntegerField()
    annee           = models.IntegerField()
    config_exercice = models.ForeignKey('parametrage.ConfigExercice', on_delete=models.PROTECT)

    # --- VERSEMENTS ---
    versement_banque   = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    versement_especes  = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    autre_versement    = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    complement_epargne = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # Mode calculé auto (BANQUE si versement_banque > 0, ESPECES si espèces > 0, sinon ECHEC)
    mode_versement = models.CharField(max_length=10, choices=MODE_CHOICES, default=ECHEC)

    # --- COTISATIONS ---
    inscription          = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    mutuelle             = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    contribution_foyer   = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    don_foyer_volontaire = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    sanction             = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    # --- PRETS ---
    pret_fonds         = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    remboursement_pret = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    mode_pret          = models.CharField(max_length=10, choices=MODE_CHOICES, blank=True)
    numero_cheque_pret = models.CharField(max_length=50, blank=True)
    retrait_partiel    = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # --- PÉNALITÉ ESPÈCES : décision ADMIN ---
    # L'admin coche cette case pour appliquer la pénalité espèces à ce mois
    # Ce n'est PAS calculé automatiquement - c'est une décision manuelle du bureau
    penalite_especes_appliquee = models.BooleanField(
        default=False,
        verbose_name="Appliquer pénalité espèces",
        help_text="Cocher pour appliquer la pénalité espèces (décision du bureau)"
    )

    # --- CHAMPS CALCULÉS (sauvegardés après saisie) ---
    bonus_malus             = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    montant_engagement      = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))
    penalite_especes_appli  = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal('0'))
    penalite_echec_appli    = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    contribution_foyer_auto = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))
    reste                   = models.DecimalField(max_digits=14, decimal_places=2, default=Decimal('0'))

    # Libellés dépenses
    libelle_depense = models.CharField(max_length=200, blank=True)
    compte_depense  = models.CharField(max_length=100, blank=True)
    autres_depenses = models.DecimalField(max_digits=12, decimal_places=2, default=Decimal('0'))

    est_valide        = models.BooleanField(default=False)
    date_saisie       = models.DateTimeField(auto_now_add=True)
    date_modification = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Tableau de bord mensuel"
        unique_together = [('adherent', 'mois', 'annee')]
        ordering = ['annee', 'mois', 'adherent__numero_ordre']

    def __str__(self):
        return f"Saisie {self.adherent.matricule} - {self.mois:02d}/{self.annee}"

    def calculer_mode_versement(self):
        """
        Excel: =IF(D>0,"BANQUE", IF(E>0,"ESPECES","ECHEC"))
        """
        if self.versement_banque > 0:
            self.mode_versement = self.BANQUE
        elif self.versement_especes > 0:
            self.mode_versement = self.ESPECES
        else:
            self.mode_versement = self.ECHEC
        return self.mode_versement

    def calculer_penalite_especes(self):
        """
        RÈGLE : La pénalité espèces est décidée MANUELLEMENT par l'admin.
        Elle est appliquée UNIQUEMENT si l'admin a coché penalite_especes_appliquee=True.
        Ce n'est PAS un calcul automatique sur le mode de versement.

        Cas d'usage réel ASELBY 2025 :
        - BAKOP paie en espèces 10 mois, mais seulement 5 000 F de pénalité totale
        - La case est cochée seulement quand le bureau le décide
        """
        if self.penalite_especes_appliquee:
            self.penalite_especes_appli = self.config_exercice.penalite_especes
        else:
            self.penalite_especes_appli = Decimal('0')
        return self.penalite_especes_appli

    def calculer_penalite_echec(self, montant_tontines_du):
        """
        Excel: =IF(I="ECHEC", (K+M+O)*20/100, 0)
        """
        cfg = self.config_exercice
        if self.mode_versement == self.ECHEC:
            self.penalite_echec_appli = montant_tontines_du * (cfg.pourcentage_penalite_echec / 100)
        else:
            self.penalite_echec_appli = Decimal('0')
        return self.penalite_echec_appli

    def calculer_contribution_foyer_auto(self, a_obtenu_lot_principal):
        """
        Excel: =IF(AC>0, 150000, 0)
        """
        cfg = self.config_exercice
        self.contribution_foyer_auto = cfg.contribution_foyer_lot_principal if a_obtenu_lot_principal else Decimal('0')
        return self.contribution_foyer_auto

    def calculer_bonus_malus(self):
        """
        Excel col C: =IF(I<>"ECHEC", D+E+F-U, 0)
        """
        if self.mode_versement != self.ECHEC:
            total_entree = self.versement_banque + self.versement_especes + self.autre_versement
            self.bonus_malus = total_entree - self.montant_engagement
        else:
            self.bonus_malus = Decimal('0')
        return self.bonus_malus
