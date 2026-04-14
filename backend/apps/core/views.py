from django.shortcuts import render
def erreur_403(request, exception=None):
    return render(request, 'errors/403.html', status=403)
def erreur_404(request, exception=None):
    return render(request, 'errors/404.html', status=404)
def erreur_500(request):
    return render(request, 'errors/500.html', status=500)
