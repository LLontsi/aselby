from functools import wraps
from django.contrib.auth.decorators import login_required
from django.shortcuts import redirect

def bureau_required(view_func):
    @wraps(view_func)
    @login_required
    def wrapper(request, *args, **kwargs):
        if not request.user.est_bureau:
            return redirect('users:mon_espace')
        return view_func(request, *args, **kwargs)
    return wrapper

def membre_required(view_func):
    @wraps(view_func)
    @login_required
    def wrapper(request, *args, **kwargs):
        if not request.user.adherent:
            return redirect('public:accueil')
        return view_func(request, *args, **kwargs)
    return wrapper
