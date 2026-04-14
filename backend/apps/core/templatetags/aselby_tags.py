from django import template
register = template.Library()

@register.filter
def fcfa(value):
    if value is None: return '0 FCFA'
    try: return f'{int(value):,} FCFA'.replace(',', ' ')
    except: return str(value)

@register.filter
def initiales(value):
    if not value: return '?'
    mots = str(value).strip().split()
    if len(mots) >= 2: return (mots[0][0] + mots[1][0]).upper()
    return mots[0][0].upper() if mots else '?'

@register.filter
def get_item(dictionary, key):
    if dictionary is None: return None
    return dictionary.get(key)