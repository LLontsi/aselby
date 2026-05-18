from django import template
register = template.Library()

@register.filter(name="get_item")
def get_item(dictionary, key):
    """{{ mon_dict|get_item:ma_cle }}"""
    if dictionary is None:
        return None
    return dictionary.get(key)