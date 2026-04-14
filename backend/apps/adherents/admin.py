from django.contrib import admin
from .models import Adherent
@admin.register(Adherent)
class AdherentAdmin(admin.ModelAdmin):
    list_display = ['matricule','nom_prenom','statut','poste_bureau','residence']
    list_filter = ['statut']
    search_fields = ['matricule','nom_prenom']
    ordering = ['numero_ordre']
