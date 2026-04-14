from django import forms
from .models import Adherent

class AdherentForm(forms.ModelForm):
    class Meta:
        model = Adherent
        fields = ['matricule','numero_ordre','nom_prenom','fonction','telephone1','telephone2','residence','date_adhesion','statut','poste_bureau']
        widgets = {
            'date_adhesion': forms.DateInput(attrs={'type':'date'}),
        }
