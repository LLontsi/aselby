from django.urls import path
from . import views

app_name = 'public'

urlpatterns = [
    # ===== PAGES PUBLIQUES =====
    path('', views.accueil, name='accueil'),
    path('a-propos/', views.apropos, name='apropos'),
    path('annonces/', views.annonces, name='annonces'),
    path('annonces/<int:pk>/', views.annonce_detail, name='annonce_detail'),
    path('activites/', views.activites, name='activites'),
    path('faq/', views.faq, name='faq'),
    path('contact/', views.contact, name='contact'),

    # ===== GESTION BUREAU — Annonces =====
    path('bureau/annonces/', views.gestion_annonces, name='gestion_annonces'),
    path('bureau/annonces/creer/', views.creer_annonce, name='creer_annonce'),
    path('bureau/annonces/<int:pk>/modifier/', views.modifier_annonce, name='modifier_annonce'),
    path('bureau/annonces/<int:pk>/supprimer/', views.supprimer_annonce, name='supprimer_annonce'),

    # ===== GESTION BUREAU — Activités =====
    path('bureau/activites/', views.gestion_activites, name='gestion_activites'),
    path('bureau/activites/creer/', views.creer_activite, name='creer_activite'),
    path('bureau/activites/<int:pk>/modifier/', views.modifier_activite, name='modifier_activite'),
    path('bureau/activites/<int:pk>/supprimer/', views.supprimer_activite, name='supprimer_activite'),

    # ===== GESTION BUREAU — FAQ =====
    path('bureau/faq/', views.gestion_faq, name='gestion_faq'),
    path('bureau/faq/creer/', views.creer_faq, name='creer_faq'),
    path('bureau/faq/<int:pk>/modifier/', views.modifier_faq, name='modifier_faq'),
    path('bureau/faq/<int:pk>/supprimer/', views.supprimer_faq, name='supprimer_faq'),
]