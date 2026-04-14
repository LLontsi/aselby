from django.shortcuts import render, get_object_or_404
from django.utils import timezone
from apps.core.mixins import bureau_required
from apps.parametrage.models import ConfigExercice
from .models import NiveauTontine, SessionTontine, ParticipationTontine

@bureau_required
def tableau_mensuel(request):
    config = ConfigExercice.get_exercice_courant()
    mois = int(request.GET.get('mois', timezone.now().month))
    annee = int(request.GET.get('annee', timezone.now().year))
    niveau_code = request.GET.get('niveau', '')  # filtre par niveau si présent

    sessions = SessionTontine.objects.filter(mois=mois, annee=annee, niveau__config_exercice=config).select_related('niveau')
    # Si aucune session pour le mois courant, afficher le dernier mois disponible
    if not sessions.exists() and not request.GET.get('mois'):
        derniere = SessionTontine.objects.filter(
            niveau__config_exercice=config
        ).order_by('-annee', '-mois').first()
        if derniere:
            mois = derniere.mois
            annee = derniere.annee
            sessions = SessionTontine.objects.filter(mois=mois, annee=annee, niveau__config_exercice=config).select_related('niveau')

    participations_qs = ParticipationTontine.objects.filter(
        session__mois=mois, session__annee=annee, session__niveau__config_exercice=config
    ).select_related('adherent', 'session__niveau').order_by('adherent__numero_ordre')

    # Filtre par niveau si paramètre ?niveau=T60
    niveau_actif = None
    if niveau_code:
        participations_qs = participations_qs.filter(session__niveau__code=niveau_code)
        niveau_actif = NiveauTontine.objects.filter(config_exercice=config, code=niveau_code).first()

    ctx = {
        'config_exercice': config,
        'sessions': sessions,
        'mois': mois,
        'annee': annee,
        'participations': participations_qs,
        'niveau_actif': niveau_actif,
        'niveau_code': niveau_code,
    }
    return render(request, 'dashboard/tontines/tableau.html', ctx)

@bureau_required
def detail_participants(request, niveau_code):
    """Vue dédiée pour voir tous les participants d'un niveau tontine sur l'année."""
    config = ConfigExercice.get_exercice_courant()
    niveau = get_object_or_404(NiveauTontine, config_exercice=config, code=niveau_code)
    # Toutes les participations de ce niveau pour l'exercice courant
    participations = ParticipationTontine.objects.filter(
        session__niveau=niveau
    ).select_related('adherent', 'session').order_by('session__mois', 'adherent__numero_ordre')
    # Sessions avec leur lot principal
    sessions = SessionTontine.objects.filter(niveau=niveau).order_by('mois')
    ctx = {
        'config_exercice': config,
        'niveau': niveau,
        'participations': participations,
        'sessions': sessions,
    }
    return render(request, 'dashboard/tontines/detail_participants.html', ctx)

@bureau_required
def enchere(request):
    config = ConfigExercice.get_exercice_courant()
    return render(request, 'dashboard/tontines/enchere.html', {'config_exercice': config})

@bureau_required
def calendrier(request):
    config = ConfigExercice.get_exercice_courant()
    niveaux = NiveauTontine.objects.filter(config_exercice=config)
    return render(request, 'dashboard/tontines/calendrier.html', {'config_exercice': config, 'niveaux': niveaux})

@bureau_required
def repartition_interets(request):
    config = ConfigExercice.get_exercice_courant()
    mois = int(request.GET.get('mois', timezone.now().month))
    annee = int(request.GET.get('annee', timezone.now().year))
    niveaux = NiveauTontine.objects.filter(config_exercice=config)
    ctx = {'config_exercice': config, 'niveaux': niveaux, 'mois': mois, 'annee': annee}
    return render(request, 'dashboard/tontines/interets.html', ctx)