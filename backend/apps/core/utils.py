from decimal import Decimal
import math

MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
           'Juillet','Août','Septembre','Octobre','Novembre','Décembre']

def rounddown(value, decimals=0):
    factor = 10 ** decimals
    return math.floor(float(value) * factor) / factor

def fcfa(value):
    if value is None: return '0 FCFA'
    return f'{int(value):,} FCFA'.replace(',', ' ')
