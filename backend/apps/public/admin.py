from django.contrib import admin
from .models import Annonce, Activite, FAQ, Contact
@admin.register(Annonce)
class AnnonceAdmin(admin.ModelAdmin):
    list_display = ['titre','date_publication','est_publiee']
    list_editable = ['est_publiee']
@admin.register(Activite)
class ActiviteAdmin(admin.ModelAdmin):
    list_display = ['titre','date','lieu','est_publiee']
    list_editable = ['est_publiee']
@admin.register(FAQ)
class FAQAdmin(admin.ModelAdmin):
    list_display = ['question','ordre','est_publiee']
    list_editable = ['ordre','est_publiee']
@admin.register(Contact)
class ContactAdmin(admin.ModelAdmin):
    list_display = ['ville','pays','telephone']
