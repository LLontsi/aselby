from django import forms
from .models import TableauDeBord


class TableauDeBordForm(forms.ModelForm):
    """
    Formulaire de saisie mensuelle du tableau de bord.

    PÉNALITÉ ESPÈCES :
    Le champ `penalite_especes_appliquee` est visible et coché par l'admin.
    Ce n'est PAS calculé automatiquement — c'est une décision manuelle du bureau.
    """
    class Meta:
        model = TableauDeBord
        exclude = [
            'mode_versement', 'bonus_malus', 'montant_engagement',
            'penalite_especes_appli', 'penalite_echec_appli',
            'contribution_foyer_auto', 'reste', 'est_valide',
            'date_saisie', 'date_modification', 'config_exercice'
        ]
        widgets = {
            'versement_banque':         forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'versement_especes':        forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'complement_epargne':       forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'pret_fonds':               forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'remboursement_pret':       forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'retrait_partiel':          forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'mutuelle':                 forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'inscription':              forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'contribution_foyer':       forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'don_foyer_volontaire':     forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            'sanction':                 forms.NumberInput(attrs={'min': 0, 'step': '1000', 'class': 'form-input'}),
            # Pénalité espèces : checkbox décision admin
            'penalite_especes_appliquee': forms.CheckboxInput(attrs={
                'class': 'form-checkbox h-4 w-4 text-red-600',
                'title': 'Cocher pour appliquer la pénalité espèces (3 000 F)'
            }),
        }
        labels = {
            'penalite_especes_appliquee': 'Appliquer pénalité espèces (3 000 F)',
        }
        help_texts = {
            'penalite_especes_appliquee': 'Décision du bureau — non automatique',
        }
