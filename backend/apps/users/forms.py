from django import forms
from django.contrib.auth.forms import AuthenticationForm
from .models import Utilisateur


class ConnexionForm(AuthenticationForm):
    """
    Connexion par nom d'utilisateur + mot de passe.
    Champ email supprimé, labels en français.
    """
    username = forms.CharField(
        label="Nom d'utilisateur",
        widget=forms.TextInput(attrs={
            'placeholder': "Votre nom d'utilisateur",
            'autocomplete': 'username',
            'autofocus': True,
        })
    )
    password = forms.CharField(
        label="Mot de passe",
        widget=forms.PasswordInput(attrs={
            'placeholder': "Votre mot de passe",
            'autocomplete': 'current-password',
        })
    )

    error_messages = {
        'invalid_login': "Nom d'utilisateur ou mot de passe incorrect.",
        'inactive'     : "Ce compte est désactivé.",
    }


class ReinitialisationMotDePasseForm(forms.Form):
    """
    Réinitialisation sans email.
    L'utilisateur s'identifie par son nom d'utilisateur + numéro de téléphone.
    """
    username = forms.CharField(
        label="Nom d'utilisateur",
        widget=forms.TextInput(attrs={'placeholder': "Votre nom d'utilisateur"})
    )
    telephone = forms.CharField(
        label="Numéro de téléphone enregistré",
        widget=forms.TextInput(attrs={'placeholder': "Ex: 655592631"})
    )

    def get_user(self):
        username  = self.cleaned_data.get('username', '').strip()
        telephone = self.cleaned_data.get('telephone', '').strip()
        try:
            user = Utilisateur.objects.get(username=username)
            adherent = user.adherent
            if adherent and (
                adherent.telephone1 == telephone
                or adherent.telephone2 == telephone
            ):
                return user
        except Utilisateur.DoesNotExist:
            pass
        return None


class DemandePretForm(forms.Form):
    """
    Formulaire de demande de prêt par un membre depuis son espace.
    """
    montant = forms.DecimalField(
        label="Montant souhaité (FCFA)",
        min_value=1,
        max_digits=12,
        decimal_places=2,
        widget=forms.NumberInput(attrs={'placeholder': "Ex: 200000"})
    )
    nombre_mois = forms.IntegerField(
        label="Durée de remboursement (mois)",
        min_value=1,
        max_value=24,
        widget=forms.NumberInput(attrs={'placeholder': "Ex: 3"})
    )
    motif = forms.CharField(
        label="Motif de la demande",
        widget=forms.Textarea(attrs={'rows': 3, 'placeholder': "Expliquez brièvement la raison de votre demande..."})
    )

    def __init__(self, *args, config=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.config = config

    def clean(self):
        cleaned = super().clean()
        montant     = cleaned.get('montant')
        nb_mois     = cleaned.get('nombre_mois')
        if montant and nb_mois and self.config:
            taux  = self.config.taux_interet_pret_mensuel
            interet = montant * (taux / 100) * nb_mois
            cleaned['interet_calcule'] = interet
            cleaned['montant_total']   = montant + interet
        return cleaned

    def save(self, commit=True):
        from apps.prets.models import Pret
        cd = self.cleaned_data
        pret = Pret(
            montant_principal = cd['montant'],
            nombre_mois       = cd['nombre_mois'],
            motif_demande     = cd['motif'],
            mode_versement    = 'BANQUE',
        )
        return pret
