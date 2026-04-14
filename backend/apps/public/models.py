from django.db import models

class Annonce(models.Model):
    titre            = models.CharField(max_length=300)
    contenu          = models.TextField()
    date_publication = models.DateField(auto_now_add=True)
    est_publiee      = models.BooleanField(default=True)
    auteur           = models.ForeignKey('users.Utilisateur', on_delete=models.SET_NULL, null=True, blank=True)
    date_creation    = models.DateTimeField(auto_now_add=True)
    class Meta:
        verbose_name = "Annonce"
        ordering = ['-date_publication']
    def __str__(self): return self.titre

class Activite(models.Model):
    titre       = models.CharField(max_length=300)
    description = models.TextField()
    date        = models.DateField()
    lieu        = models.CharField(max_length=200, blank=True)
    est_publiee = models.BooleanField(default=True)
    class Meta:
        verbose_name = "Activité"
        ordering = ['-date']
    def __str__(self): return f"{self.titre} — {self.date}"
    @property
    def est_passee(self):
        from django.utils import timezone
        return self.date < timezone.now().date()

class FAQ(models.Model):
    question    = models.CharField(max_length=500)
    reponse     = models.TextField()
    ordre       = models.IntegerField(default=0)
    est_publiee = models.BooleanField(default=True)
    class Meta:
        verbose_name = "FAQ"
        ordering = ['ordre']
    def __str__(self): return self.question

class Contact(models.Model):
    ville     = models.CharField(max_length=100, default="Yaoundé")
    pays      = models.CharField(max_length=100, default="Cameroun")
    telephone = models.CharField(max_length=20, blank=True)
    email     = models.EmailField(blank=True)
    adresse   = models.TextField(blank=True)
    horaires  = models.TextField(blank=True)
    class Meta:
        verbose_name = "Contact"
    def __str__(self): return f"Contact ASELBY — {self.ville}"
