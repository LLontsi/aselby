from django.shortcuts import render, get_object_or_404
from django.utils import timezone
from apps.parametrage.models import ConfigExercice
from apps.adherents.models import Adherent
from apps.public.models import Annonce, Activite, FAQ, Contact


def _get_stats():
    config = ConfigExercice.get_exercice_courant()
    return {
        'nb_adherents'       : Adherent.objects.filter(statut='ACTIF').count(),
        'nb_niveaux_tontine' : 3,
        'annee_fondation'    : 2016,
        'annee_courante'     : config.annee if config else timezone.now().year,
        'nb_reunions_annee'  : 12,
    }

def _ctx_base():
    return {
        'stats'           : _get_stats(),
        'config_exercice' : ConfigExercice.get_exercice_courant(),
    }

def accueil(request):
    ctx = _ctx_base()
    ctx['annonces_recentes'] = Annonce.objects.filter(est_publiee=True).order_by('-date_publication')[:3]
    ctx['prochaine_activite'] = Activite.objects.filter(date__gte=timezone.now().date(), est_publiee=True).order_by('date').first()
    return render(request, 'public/accueil.html', ctx)

def apropos(request):
    ctx = _ctx_base()
    POSTES_GAUCHE = ['Président', 'Secrétaire Général', 'Censeur']
    POSTES_DROITE = ['Vice-Président', 'Trésorier', 'Commissaire aux Comptes']
    bureau = list(Adherent.objects.filter(statut='ACTIF', poste_bureau__isnull=False).exclude(poste_bureau='').order_by('numero_ordre'))
    ctx.update({
        'membres_bureau_gauche': [m for m in bureau if m.poste_bureau in POSTES_GAUCHE],
        'membres_bureau_droite': [m for m in bureau if m.poste_bureau in POSTES_DROITE],
        'postes_gauche': POSTES_GAUCHE,
        'postes_droite': POSTES_DROITE,
    })
    return render(request, 'public/apropos.html', ctx)

def annonces(request):
    from django.core.paginator import Paginator
    ctx = _ctx_base()
    qs = Annonce.objects.filter(est_publiee=True).order_by('-date_publication')
    paginator = Paginator(qs, 9)
    page_obj = paginator.get_page(request.GET.get('page'))
    ctx['page_obj'] = page_obj
    ctx['annonces'] = page_obj.object_list
    return render(request, 'public/annonces.html', ctx)

def annonce_detail(request, pk):
    ctx = _ctx_base()
    ctx['annonce'] = get_object_or_404(Annonce, pk=pk, est_publiee=True)
    ctx['annonces_liees'] = Annonce.objects.filter(est_publiee=True).exclude(pk=pk).order_by('-date_publication')[:3]
    return render(request, 'public/annonce_detail.html', ctx)

def activites(request):
    ctx = _ctx_base()
    aujourd_hui = timezone.now().date()
    ctx['activites_avenir']  = Activite.objects.filter(date__gte=aujourd_hui, est_publiee=True).order_by('date')
    ctx['activites_passees'] = Activite.objects.filter(date__lt=aujourd_hui, est_publiee=True).order_by('-date')[:10]
    return render(request, 'public/activites.html', ctx)

def faq(request):
    ctx = _ctx_base()
    ctx['faqs'] = FAQ.objects.filter(est_publiee=True).order_by('ordre')
    return render(request, 'public/faq.html', ctx)

def contact(request):
    ctx = _ctx_base()
    ctx['infos_contact'] = Contact.objects.first()
    return render(request, 'public/contact.html', ctx)


"""
Views admin pour Annonces, Activités, FAQ
À AJOUTER à la fin de backend/apps/public/views.py
"""
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib import messages
from django.utils import timezone


def _config():
    from apps.parametrage.models import ConfigExercice
    return ConfigExercice.get_exercice_courant()


# ===== ANNONCES ADMIN =====
def gestion_annonces(request):
    from apps.core.mixins import bureau_required
    from apps.public.models import Annonce
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    annonces = Annonce.objects.all().order_by('-date_publication')
    return render(request, 'dashboard/public_admin/gestion_annonces.html', {'annonces': annonces, 'config_exercice': _config()})


def creer_annonce(request):
    from apps.public.models import Annonce
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    if request.method == 'POST':
        titre = request.POST.get('titre', '').strip()
        contenu = request.POST.get('contenu', '').strip()
        if titre and contenu:
            Annonce.objects.create(
                titre=titre, contenu=contenu,
                est_publiee=bool(request.POST.get('est_publiee')),
                date_publication=request.POST.get('date_publication') or timezone.now().date(),
            )
            messages.success(request, 'Annonce créée avec succès.')
            return redirect('public:gestion_annonces')
        messages.error(request, 'Titre et contenu obligatoires.')
    return render(request, 'dashboard/public_admin/form_annonce.html', {'config_exercice': _config()})


def modifier_annonce(request, pk):
    from apps.public.models import Annonce
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    annonce = get_object_or_404(Annonce, pk=pk)
    if request.method == 'POST':
        annonce.titre = request.POST.get('titre', annonce.titre).strip()
        annonce.contenu = request.POST.get('contenu', annonce.contenu).strip()
        annonce.est_publiee = bool(request.POST.get('est_publiee'))
        if request.POST.get('date_publication'):
            annonce.date_publication = request.POST.get('date_publication')
        annonce.save()
        messages.success(request, 'Annonce mise à jour.')
        return redirect('public:gestion_annonces')
    return render(request, 'dashboard/public_admin/form_annonce.html', {'annonce': annonce, 'config_exercice': _config()})


def supprimer_annonce(request, pk):
    from apps.public.models import Annonce
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    annonce = get_object_or_404(Annonce, pk=pk)
    if request.method == 'POST':
        annonce.delete()
        messages.success(request, 'Annonce supprimée.')
    return redirect('public:gestion_annonces')


# ===== ACTIVITÉS ADMIN =====
def gestion_activites(request):
    from apps.public.models import Activite
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    return render(request, 'dashboard/public_admin/gestion_activites.html', {'activites': Activite.objects.all().order_by('-date'), 'config_exercice': _config()})


def creer_activite(request):
    from apps.public.models import Activite
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    if request.method == 'POST':
        titre = request.POST.get('titre', '').strip()
        date  = request.POST.get('date')
        if titre and date:
            Activite.objects.create(
                titre=titre, description=request.POST.get('description', '').strip(),
                date=date, lieu=request.POST.get('lieu', '').strip(),
                est_publiee=bool(request.POST.get('est_publiee')),
            )
            messages.success(request, 'Activité créée.')
            return redirect('public:gestion_activites')
        messages.error(request, 'Titre et date obligatoires.')
    return render(request, 'dashboard/public_admin/form_activite.html', {'config_exercice': _config()})


def modifier_activite(request, pk):
    from apps.public.models import Activite
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    activite = get_object_or_404(Activite, pk=pk)
    if request.method == 'POST':
        activite.titre = request.POST.get('titre', activite.titre).strip()
        activite.description = request.POST.get('description', activite.description).strip()
        activite.lieu = request.POST.get('lieu', activite.lieu).strip()
        activite.est_publiee = bool(request.POST.get('est_publiee'))
        if request.POST.get('date'):
            activite.date = request.POST.get('date')
        activite.save()
        messages.success(request, 'Activité mise à jour.')
        return redirect('public:gestion_activites')
    return render(request, 'dashboard/public_admin/form_activite.html', {'activite': activite, 'config_exercice': _config()})


def supprimer_activite(request, pk):
    from apps.public.models import Activite
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    activite = get_object_or_404(Activite, pk=pk)
    if request.method == 'POST':
        activite.delete()
        messages.success(request, 'Activité supprimée.')
    return redirect('public:gestion_activites')


# ===== FAQ ADMIN =====
def gestion_faq(request):
    from apps.public.models import FAQ
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    return render(request, 'dashboard/public_admin/gestion_faq.html', {'faqs': FAQ.objects.all().order_by('ordre'), 'config_exercice': _config()})


def creer_faq(request):
    from apps.public.models import FAQ
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    if request.method == 'POST':
        question = request.POST.get('question', '').strip()
        reponse  = request.POST.get('reponse', '').strip()
        if question and reponse:
            FAQ.objects.create(
                question=question, reponse=reponse,
                ordre=int(request.POST.get('ordre', 0) or 0),
                est_publiee=bool(request.POST.get('est_publiee')),
            )
            messages.success(request, 'Question FAQ créée.')
            return redirect('public:gestion_faq')
        messages.error(request, 'Question et réponse obligatoires.')
    return render(request, 'dashboard/public_admin/form_faq.html', {'config_exercice': _config()})


def modifier_faq(request, pk):
    from apps.public.models import FAQ
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    faq = get_object_or_404(FAQ, pk=pk)
    if request.method == 'POST':
        faq.question = request.POST.get('question', faq.question).strip()
        faq.reponse  = request.POST.get('reponse', faq.reponse).strip()
        faq.ordre    = int(request.POST.get('ordre', faq.ordre) or 0)
        faq.est_publiee = bool(request.POST.get('est_publiee'))
        faq.save()
        messages.success(request, 'Question FAQ mise à jour.')
        return redirect('public:gestion_faq')
    return render(request, 'dashboard/public_admin/form_faq.html', {'faq': faq, 'config_exercice': _config()})


def supprimer_faq(request, pk):
    from apps.public.models import FAQ
    if not request.user.is_authenticated or not request.user.est_bureau:
        return redirect('users:connexion')
    faq = get_object_or_404(FAQ, pk=pk)
    if request.method == 'POST':
        faq.delete()
        messages.success(request, 'Question FAQ supprimée.')
    return redirect('public:gestion_faq')