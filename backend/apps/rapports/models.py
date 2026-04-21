from django.db import models
from decimal import Decimal
D = Decimal


class ComplementMouvement(models.Model):
    """
    Complément AUTREMOUVEMENT par adhérent/mois.
    Toutes les données déjà en BD sont référencées par FK.
    Seuls les champs vraiment nouveaux (non présents ailleurs) sont stockés ici.
    """
    # ── Clés vers les données existantes ──────────────────────────
    adherent        = models.ForeignKey(
        'adherents.Adherent', on_delete=models.PROTECT,
        related_name='complements_mouvement')
    mouvement_fonds = models.OneToOneField(
        'fonds.MouvementFonds', on_delete=models.PROTECT,
        related_name='complement',
        help_text="Lie ce complément au MouvementFonds existant (même adhérent/mois/annee)")
    tableau_bord    = models.OneToOneField(
        'saisie.TableauDeBord', on_delete=models.PROTECT,
        related_name='complement_mouvement',
        help_text="Lie ce complément au TableauDeBord existant")
    config_exercice = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)

    # ── Champs nouveaux : prêts / chèques ─────────────────────────
    # Ces infos ne sont stockées nulle part ailleurs dans la BD
    interet_pret_fonds           = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    numero_cheque_versement      = models.CharField(max_length=50, blank=True,
        help_text="N° chèque du versement mensuel (col 20 AUTREMOUVEMENT)")
    remboursement_pret_fonds     = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Montant remboursé ce mois (col 21)")
    mode_versement_remboursement = models.CharField(max_length=20, blank=True,
        help_text="BANQUE/ESPECES/ECHEC (col 22)")
    mode_versement_pret          = models.CharField(max_length=20, blank=True,
        help_text="Mode de versement du prêt (col 23)")
    numero_cheque_pret           = models.CharField(max_length=50, blank=True,
        help_text="N° chèque du prêt (col 24)")
    penalite_pret_fonds          = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    penalite_fonds               = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))

    # ── Champs nouveaux : pénalités tontine ───────────────────────
    # NB : penalite_echec_appli est dans TableauDeBord,
    #      penalite_retard_tontine n'est pas stocké séparément
    penalite_echec_tontine  = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    penalite_retard_tontine = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    remboursement_transport = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))

    # ── Champs nouveaux : dépenses détail ─────────────────────────
    # Les modèles DepenseFondsRoulement etc. stockent le total global,
    # mais pas la part individuelle par adhérent
    depense_fonds_roulement    = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    depense_frais_exceptionnel = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    depense_fonds_mutuelle     = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    depense_collation          = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    depense_penalite_banque    = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))

    # ── Champs résumé annuel uniquement ───────────────────────────
    pret_definitif     = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'))
    date_remboursement = models.CharField(max_length=30, blank=True)
    statut_pret        = models.CharField(max_length=20, blank=True)
    don_volontaire     = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))

    date_saisie       = models.DateTimeField(auto_now_add=True)
    date_modification = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Complément Mouvement"
        # L'unicité est garantie par le OneToOne vers MouvementFonds
        ordering = ['mouvement_fonds__annee', 'mouvement_fonds__mois',
                    'adherent__numero_ordre']

    def __str__(self):
        return (f"MVT {self.adherent.matricule} "
                f"{self.mouvement_fonds.mois:02d}/{self.mouvement_fonds.annee}")

    # ── Propriétés : accès transparent aux données FK ─────────────
    @property
    def mois(self):
        return self.mouvement_fonds.mois

    @property
    def annee(self):
        return self.mouvement_fonds.annee

    # Données depuis MouvementFonds (lecture seule via FK)
    @property
    def fonds_depart(self):
        return self.mouvement_fonds.capital_compose_precedent

    @property
    def reconduction(self):
        return self.mouvement_fonds.reconduction

    @property
    def retrait_partiel(self):
        return self.mouvement_fonds.retrait_partiel

    @property
    def fonds_definitif(self):
        return self.mouvement_fonds.fonds_definitif

    @property
    def base_calcul_interet(self):
        return self.mouvement_fonds.base_calcul_interet

    @property
    def repartition_interet(self):
        return self.mouvement_fonds.interet_attribue

    @property
    def capital_compose(self):
        return self.mouvement_fonds.capital_compose

    @property
    def sanction(self):
        return self.mouvement_fonds.sanction

    @property
    def reste(self):
        return self.mouvement_fonds.reste

    @property
    def epargne(self):
        return self.mouvement_fonds.epargne_nette

    @property
    def fonds_roulement(self):
        return self.mouvement_fonds.fonds_roulement

    @property
    def frais_exceptionnel(self):
        return self.mouvement_fonds.frais_exceptionnels

    @property
    def collation(self):
        return self.mouvement_fonds.collation

    # Données depuis TableauDeBord (lecture seule via FK)
    @property
    def penalite_vst_especes(self):
        return self.tableau_bord.penalite_especes_appli

    @property
    def inscription(self):
        return self.tableau_bord.inscription

    @property
    def mutuelle(self):
        return self.tableau_bord.mutuelle

    @property
    def pret_fonds(self):
        return self.tableau_bord.pret_fonds

    @property
    def contribution_foyer(self):
        return self.tableau_bord.contribution_foyer

    @property
    def autres_depenses(self):
        return self.tableau_bord.autres_depenses


class ComplementHistorique(models.Model):
    """
    Complément TABBORD/HISTORIQUE par adhérent/mois.
    Toutes les données déjà en BD sont référencées par FK.
    Seuls les champs vraiment nouveaux sont stockés ici.
    """
    # ── Clés vers les données existantes ──────────────────────────
    adherent     = models.ForeignKey(
        'adherents.Adherent', on_delete=models.PROTECT,
        related_name='complements_historique')
    tableau_bord = models.OneToOneField(
        'saisie.TableauDeBord', on_delete=models.PROTECT,
        related_name='complement_historique',
        help_text="Lie ce complément au TableauDeBord existant")
    config_exercice = models.ForeignKey(
        'parametrage.ConfigExercice', on_delete=models.PROTECT)

    # ── Champs nouveaux : modes et chèques ────────────────────────
    # ParticipationTontine stocke montant/parts/pénalités mais pas ces modes
    mode_paiement_tontine  = models.CharField(max_length=20, blank=True,
        help_text="BANQUE/ESPECES (col 08 TABBORD)")
    num_cheque_versement   = models.CharField(max_length=50, blank=True,
        help_text="N° chèque versement tontine (col 27)")
    mode_paiement_lot      = models.CharField(max_length=20, blank=True,
        help_text="Mode paiement lot principal (col 31)")
    num_cheque_lot         = models.CharField(max_length=50, blank=True,
        help_text="N° chèque lot principal (col 32)")
    num_cheque_petit_lot   = models.CharField(max_length=50, blank=True,
        help_text="N° chèque petit lot (col 40)")
    mode_remb_petit_lot    = models.CharField(max_length=20, blank=True,
        help_text="Mode remboursement petit lot (col 41)")
    autres_mode_verst      = models.CharField(max_length=100, blank=True,
        help_text="Autres modes de versement (col 54)")
    numero_cheque_effectif = models.CharField(max_length=50, blank=True,
        help_text="N° chèque effectif (col 56)")

    # ── Champs nouveaux : financiers ──────────────────────────────
    epargne_assurance      = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Épargne assurance (col 23)")
    montant_cheque_effectif = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'),
        help_text="Montant chèque effectif (col 55)")
    interet_pret           = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Intérêt prêt ce mois (col 57)")
    nbre_mois_pret         = models.IntegerField(default=0,
        help_text="Nombre de mois du prêt (col 58)")
    montant_depense        = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'),
        help_text="Montant dépense individuelle (col 68)")
    penalite_pret_fonds    = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))
    remboursement_transport = models.DecimalField(max_digits=12, decimal_places=2, default=D('0'))

    # ── Champs résumé annuel uniquement ───────────────────────────
    pret_definitif     = models.DecimalField(max_digits=14, decimal_places=2, default=D('0'))
    nombre_mois_pret   = models.IntegerField(default=0)
    date_remboursement = models.CharField(max_length=30, blank=True)
    statut_pret        = models.CharField(max_length=20, blank=True)

    date_saisie       = models.DateTimeField(auto_now_add=True)
    date_modification = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Complément Historique"
        ordering = ['tableau_bord__annee', 'tableau_bord__mois',
                    'adherent__numero_ordre']

    def __str__(self):
        return (f"HISTO {self.adherent.matricule} "
                f"{self.tableau_bord.mois:02d}/{self.tableau_bord.annee}")

    @property
    def mois(self):
        return self.tableau_bord.mois

    @property
    def annee(self):
        return self.tableau_bord.annee

    # ── Données depuis TableauDeBord (via FK) ─────────────────────
    @property
    def bonus_malus(self):
        return self.tableau_bord.bonus_malus

    @property
    def versement_banque(self):
        return self.tableau_bord.versement_banque

    @property
    def versement_especes(self):
        return self.tableau_bord.versement_especes

    @property
    def autre_versement(self):
        return self.tableau_bord.autre_versement

    @property
    def complement_epargne(self):
        return self.tableau_bord.complement_epargne

    @property
    def penalite_vst_especes(self):
        return self.tableau_bord.penalite_especes_appli

    @property
    def montant_engagement(self):
        return self.tableau_bord.montant_engagement

    @property
    def sanction(self):
        return self.tableau_bord.sanction

    @property
    def inscription(self):
        return self.tableau_bord.inscription

    @property
    def pret_fonds(self):
        return self.tableau_bord.pret_fonds

    @property
    def remboursement_pret(self):
        return self.tableau_bord.remboursement_pret

    @property
    def reste(self):
        return self.tableau_bord.reste

    @property
    def mutuelle(self):
        return self.tableau_bord.mutuelle

    @property
    def retrait_partiel(self):
        return self.tableau_bord.retrait_partiel

    @property
    def montant_cheque(self):
        return self.tableau_bord.versement_banque

    @property
    def montant_especes(self):
        return self.tableau_bord.versement_especes

    @property
    def contribution_foyer(self):
        return self.tableau_bord.contribution_foyer

    @property
    def autres_depenses(self):
        return self.tableau_bord.autres_depenses

    @property
    def don_foyer_volontaire(self):
        return self.tableau_bord.don_foyer_volontaire

    @property
    def libelle_depense(self):
        return self.tableau_bord.libelle_depense

    @property
    def compte_depense(self):
        return self.tableau_bord.compte_depense

    # ── Données depuis ParticipationTontine (via FK) ──────────────
    def _participation(self, niveau_code):
        """Retourne la ParticipationTontine de cet adhérent pour ce mois/niveau."""
        from apps.tontines.models import ParticipationTontine
        return ParticipationTontine.objects.filter(
            adherent=self.adherent,
            session__niveau__code=niveau_code,
            session__mois=self.mois,
            session__annee=self.annee,
        ).first()

    @property
    def nbre_lot_t35(self):
        p = self._participation('T35')
        return p.nombre_parts if p else 0

    @property
    def tontine_35(self):
        p = self._participation('T35')
        return p.montant_verse if p else D('0')

    @property
    def nbre_lot_t75(self):
        p = self._participation('T75')
        return p.nombre_parts if p else 0

    @property
    def tontine_75(self):
        p = self._participation('T75')
        return p.montant_verse if p else D('0')

    @property
    def nbre_lot_t100(self):
        p = self._participation('T100')
        return p.nombre_parts if p else 0

    @property
    def tontine_100(self):
        p = self._participation('T100')
        return p.montant_verse if p else D('0')

    @property
    def tontine_mois(self):
        return self.tontine_35 + self.tontine_75 + self.tontine_100

    @property
    def achat_lot_t35(self):
        p = self._participation('T35')
        return p.montant_lot_principal if p else D('0')

    @property
    def achat_lot_t75(self):
        p = self._participation('T75')
        return p.montant_lot_principal if p else D('0')

    @property
    def achat_lot_t100(self):
        p = self._participation('T100')
        return p.montant_lot_principal if p else D('0')

    @property
    def vente_petit_lot_t35(self):
        p = self._participation('T35')
        return p.vente_petit_lot if p else D('0')

    @property
    def vente_petit_lot_t75(self):
        p = self._participation('T75')
        return p.vente_petit_lot if p else D('0')

    @property
    def vente_petit_lot_t100(self):
        p = self._participation('T100')
        return p.vente_petit_lot if p else D('0')

    @property
    def interet_petit_lot_t35(self):
        p = self._participation('T35')
        return p.interet_petit_lot if p else D('0')

    @property
    def interet_petit_lot_t75(self):
        p = self._participation('T75')
        return p.interet_petit_lot if p else D('0')

    @property
    def interet_petit_lot_t100(self):
        p = self._participation('T100')
        return p.interet_petit_lot if p else D('0')

    @property
    def remb_petit_lot_t35(self):
        p = self._participation('T35')
        return p.remboursement_petit_lot if p else D('0')

    @property
    def remb_petit_lot_t75(self):
        p = self._participation('T75')
        return p.remboursement_petit_lot if p else D('0')

    @property
    def remb_petit_lot_t100(self):
        p = self._participation('T100')
        return p.remboursement_petit_lot if p else D('0')

    @property
    def penalite_retard_tontine(self):
        """Somme des pénalités retard tontine sur toutes les participations du mois."""
        from apps.tontines.models import ParticipationTontine
        from django.db.models import Sum
        total = ParticipationTontine.objects.filter(
            adherent=self.adherent,
            session__mois=self.mois,
            session__annee=self.annee,
        ).aggregate(s=Sum('penalite_especes'))['s']
        return total or D('0')

    @property
    def penalite_echec_tontine(self):
        from apps.tontines.models import ParticipationTontine
        from django.db.models import Sum
        total = ParticipationTontine.objects.filter(
            adherent=self.adherent,
            session__mois=self.mois,
            session__annee=self.annee,
        ).aggregate(s=Sum('penalite_echec'))['s']
        return total or D('0')

    @property
    def montant_t25(self):
        """Alias TABBORD col 48 = même chose que tontine_35."""
        return self.tontine_35

    @property
    def montant_t75(self):
        return self.tontine_75

    @property
    def montant_t100(self):
        return self.tontine_100

    @property
    def depense_fonds_roulement(self):
        """Part individuelle = depuis TableauDeBord (libelle FONDS DE ROULEMENT)."""
        if self.tableau_bord.compte_depense == 'FONDS DE ROULEMENT':
            return self.tableau_bord.autres_depenses
        return D('0')

    @property
    def depense_frais_excep(self):
        if self.tableau_bord.compte_depense == 'FRAIS EXCEPTIONNELS':
            return self.tableau_bord.autres_depenses
        return D('0')

    @property
    def depense_fonds_mutuel(self):
        if self.tableau_bord.compte_depense == 'FONDS MUTUELLE':
            return self.tableau_bord.autres_depenses
        return D('0')

    @property
    def depense_collation(self):
        if self.tableau_bord.compte_depense == 'COLLATION':
            return self.tableau_bord.autres_depenses
        return D('0')

    @property
    def depense_penalite_banque(self):
        if self.tableau_bord.compte_depense == 'PENALITE BANQUE':
            return self.tableau_bord.autres_depenses
        return D('0')

    @property
    def mode_paiement_remb_pret(self):
        return self.tableau_bord.mode_versement if hasattr(
            self.tableau_bord, 'mode_versement') else ''

    @property
    def mode_paiement_pret(self):
        """Alias — même valeur que mode_paiement_remb_pret."""
        return self.mode_paiement_remb_pret