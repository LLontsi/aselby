from apps.parametrage.models import ConfigExercice
def config_exercice(request):
    return {'config_exercice': ConfigExercice.get_exercice_courant()}
