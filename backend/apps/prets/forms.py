from django import forms
from .models import Pret, RemboursementPret

class PretForm(forms.ModelForm):
    class Meta:
        model = Pret
        fields = ['montant_principal','nombre_mois','date_octroi','date_echeance','mode_versement','numero_cheque']
        widgets = {
            'date_octroi': forms.DateInput(attrs={'type':'date'}),
            'date_echeance': forms.DateInput(attrs={'type':'date'}),
        }

class RemboursementForm(forms.ModelForm):
    class Meta:
        model = RemboursementPret
        fields = ['mois','annee','montant','mode_versement','numero_cheque']
