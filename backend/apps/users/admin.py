from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from .models import Utilisateur
@admin.register(Utilisateur)
class UtilisateurAdmin(UserAdmin):
    list_display = ['username','nom_complet','role','est_actif']
    list_filter = ['role','est_actif']
    search_fields = ['username','nom_complet']
    fieldsets = (
        (None, {'fields': ('username','password')}),
        ('Infos', {'fields': ('nom_complet','role','adherent')}),
        ('Permissions', {'fields': ('est_actif','is_staff','is_superuser')}),
    )
    add_fieldsets = ((None, {'fields': ('username','nom_complet','role','password1','password2')}),)
    ordering = ['nom_complet']
