from django import forms
from .models import (ConfigExercice, DepenseFondsRoulement,
                     DepenseFraisExceptionnels, DepenseCollation,
                     DepenseFoyer, AgioBancaire, AutreDepense)

INPUT_NUM = {'class': 'f-input', 'min': '0'}
INPUT_PCT = {'class': 'f-input', 'min': '0', 'max': '100', 'step': '0.01'}


class ConfigExerciceForm(forms.ModelForm):
    class Meta:
        model = ConfigExercice
        exclude = ['est_ouvert', 'date_ouverture', 'date_cloture']
        widgets = {
            'annee': forms.NumberInput(attrs={'class': 'f-input', 'min': '2020', 'max': '2050'}),
            # Tontines
            'taux_t35':            forms.NumberInput(attrs=INPUT_NUM),
            'versement_t35':       forms.NumberInput(attrs=INPUT_NUM),
            'taux_t75':            forms.NumberInput(attrs=INPUT_NUM),
            'taux_t100':           forms.NumberInput(attrs=INPUT_NUM),
            'diviseur_interet_t35': forms.NumberInput(attrs={'class': 'f-input', 'min': '1'}),
            # Fonds
            'seuil_eligibilite_interets':  forms.NumberInput(attrs=INPUT_NUM),
            'fonds_roulement_mensuel':     forms.NumberInput(attrs=INPUT_NUM),
            'frais_exceptionnels_mensuel': forms.NumberInput(attrs=INPUT_NUM),
            'collation_mensuelle':         forms.NumberInput(attrs=INPUT_NUM),
            'mutuelle_mensuelle':          forms.NumberInput(attrs=INPUT_NUM),
            'montant_inscription':         forms.NumberInput(attrs=INPUT_NUM),
            'complement_fonds_fin_exercice': forms.NumberInput(attrs=INPUT_NUM),
            # Pénalités
            'penalite_especes':              forms.NumberInput(attrs=INPUT_NUM),
            'pourcentage_penalite_echec':    forms.NumberInput(attrs=INPUT_PCT),
            'nb_echecs_max_avant_liste_rouge': forms.NumberInput(attrs={'class': 'f-input', 'min': '1'}),
            # Prêts
            'taux_interet_pret_mensuel': forms.NumberInput(attrs=INPUT_PCT),
            'majoration_retard_mois_1':  forms.NumberInput(attrs=INPUT_PCT),
            'majoration_retard_mois_2':  forms.NumberInput(attrs=INPUT_PCT),
            'montant_min_pret':          forms.NumberInput(attrs=INPUT_NUM),
            'montant_max_pret':          forms.NumberInput(attrs=INPUT_NUM),
            # Foyer / Mutuelle
            'contribution_foyer_lot_principal':  forms.NumberInput(attrs=INPUT_NUM),
            'complement_mutuelle_fin_exercice':  forms.NumberInput(attrs=INPUT_NUM),
        }

    def clean(self):
        cleaned = super().clean()
        pk = self.instance.pk if self.instance else None
        if ConfigExercice.objects.filter(est_ouvert=True).exclude(pk=pk).exists():
            raise forms.ValidationError(
                "Un exercice est déjà ouvert. Clôturez-le avant d'en ouvrir un nouveau.")
        return cleaned


class DepenseFondsRoulementForm(forms.ModelForm):
    class Meta:
        model = DepenseFondsRoulement
        fields = ['date', 'libelle', 'montant']
        widgets = {
            'date':    forms.DateInput(attrs={'class': 'f-input', 'type': 'date'}),
            'libelle': forms.TextInput(attrs={'class': 'f-input',
                'placeholder': 'Ex: TAXI BANQUE, COMMISSION, PRODUCTION RAPPORT...'}),
            'montant': forms.NumberInput(attrs={'class': 'f-input', 'min': '0', 'step': '100'}),
        }


class DepenseFraisExceptionnelsForm(forms.ModelForm):
    class Meta:
        model = DepenseFraisExceptionnels
        fields = ['date', 'libelle', 'montant']
        widgets = {
            'date':    forms.DateInput(attrs={'class': 'f-input', 'type': 'date'}),
            'libelle': forms.TextInput(attrs={'class': 'f-input',
                'placeholder': 'Ex: GRATIFICATION 2025, REMBOURSEMENT...'}),
            'montant': forms.NumberInput(attrs={'class': 'f-input', 'min': '0', 'step': '100'}),
        }


class DepenseCollationForm(forms.ModelForm):
    class Meta:
        model = DepenseCollation
        fields = ['date', 'libelle', 'montant']
        widgets = {
            'date':    forms.DateInput(attrs={'class': 'f-input', 'type': 'date'}),
            'libelle': forms.TextInput(attrs={'class': 'f-input',
                'placeholder': 'Ex: COLLATION SEANCE JANVIER'}),
            'montant': forms.NumberInput(attrs={'class': 'f-input', 'min': '0', 'step': '100'}),
        }


class DepenseFoyerForm(forms.ModelForm):
    class Meta:
        model = DepenseFoyer
        fields = ['date', 'libelle', 'montant']
        widgets = {
            'date':    forms.DateInput(attrs={'class': 'f-input', 'type': 'date'}),
            'libelle': forms.TextInput(attrs={'class': 'f-input',
                'placeholder': 'Ex: VERSE DANS LE COMPTE CONSEIL'}),
            'montant': forms.NumberInput(attrs={'class': 'f-input', 'min': '0', 'step': '1000'}),
        }


class AgioBancaireForm(forms.ModelForm):
    class Meta:
        model = AgioBancaire
        fields = ['mois', 'annee', 'montant_agio', 'interet_crediteur']
        widgets = {
            'mois':  forms.NumberInput(attrs={'class': 'f-input', 'min': '1', 'max': '12'}),
            'annee': forms.NumberInput(attrs={'class': 'f-input', 'min': '2020'}),
            'montant_agio':      forms.NumberInput(attrs={'class': 'f-input', 'min': '0', 'step': '1'}),
            'interet_crediteur': forms.NumberInput(attrs={'class': 'f-input', 'min': '0', 'step': '1'}),
        }


class AutreDepenseForm(forms.ModelForm):
    class Meta:
        model = AutreDepense
        fields = ['date', 'libelle', 'montant']
        widgets = {
            'date':    forms.DateInput(attrs={'class': 'f-input', 'type': 'date'}),
            'libelle': forms.TextInput(attrs={'class': 'f-input',
                'placeholder': 'Ex: SOLDE TERRAIN, AVANCE NOTAIRE, COMMISSION...'}),
            'montant': forms.NumberInput(attrs={'class': 'f-input', 'min': '0', 'step': '100'}),
        }
