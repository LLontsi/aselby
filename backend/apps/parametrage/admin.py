from django.contrib import admin
from .models import ConfigExercice
@admin.register(ConfigExercice)
class ConfigExerciceAdmin(admin.ModelAdmin):
    list_display = ['annee','est_ouvert','taux_t35','taux_t75','taux_t100']
    readonly_fields = ['date_ouverture','date_cloture']
