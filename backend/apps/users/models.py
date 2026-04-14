from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.db import models


class UtilisateurManager(BaseUserManager):
    def create_user(self, username, password=None, **extra_fields):
        if not username:
            raise ValueError("Le nom d'utilisateur est obligatoire")
        user = self.model(username=username.strip(), **extra_fields)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_superuser(self, username, password=None, **extra_fields):
        extra_fields.setdefault('role', 'BUREAU')
        extra_fields.setdefault('is_staff', True)
        extra_fields.setdefault('is_superuser', True)
        return self.create_user(username, password, **extra_fields)


class Utilisateur(AbstractBaseUser, PermissionsMixin):
    BUREAU = 'BUREAU'
    MEMBRE = 'MEMBRE'
    ROLE_CHOICES = [(BUREAU, 'Bureau'), (MEMBRE, 'Membre')]

    username      = models.CharField(max_length=100, unique=True)
    nom_complet   = models.CharField(max_length=200)
    role          = models.CharField(max_length=10, choices=ROLE_CHOICES, default=MEMBRE)
    adherent      = models.OneToOneField(
        'adherents.Adherent', on_delete=models.SET_NULL,
        null=True, blank=True, related_name='compte_utilisateur'
    )
    est_actif     = models.BooleanField(default=True)
    is_staff      = models.BooleanField(default=False)
    date_creation = models.DateTimeField(auto_now_add=True)

    objects = UtilisateurManager()
    USERNAME_FIELD  = 'username'
    REQUIRED_FIELDS = ['nom_complet']

    class Meta:
        verbose_name = "Utilisateur"
        ordering = ['nom_complet']

    def __str__(self):
        return f"{self.nom_complet} ({self.get_role_display()})"

    @property
    def est_bureau(self):
        return self.role == self.BUREAU

    @property
    def is_active(self):
        return self.est_actif
