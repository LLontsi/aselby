from django import template
register = template.Library()

@register.filter
def get_attr(obj, attr):
    """Récupère un attribut d'un objet par nom."""
    try:
        val = getattr(obj, attr, '')
        if val is None:
            return '0'
        return val
    except Exception:
        return '0'

@register.filter
def split(value, sep):
    return value.split(sep)

@register.filter  
def index(lst, i):
    try:
        return lst[int(i)]
    except Exception:
        return ''

@register.filter
def get_mois(mois_list, index):
    """Retourne le nom du mois depuis la liste MOIS_FR. Usage: mois_fr|get_mois:4 → 'Avril'"""
    try:
        return mois_list[int(index)]
    except (IndexError, TypeError, ValueError):
        return ''