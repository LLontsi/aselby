from django.db import migrations, models

class Migration(migrations.Migration):
    dependencies = [
        ('saisie', '0001_initial'),
    ]
    operations = [
        # Champs montant_lot (ancienne migration 0002)
        migrations.AddField(
            model_name='saisiemonthly',
            name='montant_lot_t75',
            field=models.DecimalField(
                max_digits=12, decimal_places=2, default=0,
                verbose_name="Montant lot principal T75 reçu",
                help_text="Col AX TABBORD: T75.montant_lot_principal reçu ce mois"
            ),
        ),
        migrations.AddField(
            model_name='saisiemonthly',
            name='montant_lot_t100',
            field=models.DecimalField(
                max_digits=12, decimal_places=2, default=0,
                verbose_name="Montant lot principal T100 reçu",
                help_text="Col AY TABBORD: T100.montant_lot_principal reçu ce mois"
            ),
        ),
        # Champs saisie manuelle pénalité et intérêt
        migrations.AddField(
            model_name='saisiemonthly',
            name='penalite_versement_especes_saisi',
            field=models.DecimalField(
                max_digits=12, decimal_places=2, default=0,
                verbose_name="Pénalité versement espèces (saisie manuelle)",
                help_text="Si 0: calculé auto depuis config. Sinon: valeur saisie."
            ),
        ),
        migrations.AddField(
            model_name='saisiemonthly',
            name='interet_pret_saisi',
            field=models.DecimalField(
                max_digits=12, decimal_places=2, default=0,
                verbose_name="Intérêt prêt fonds (saisi manuellement)",
                help_text="Si 0: calculé auto (taux×remb×nb_mois). Sinon: valeur saisie."
            ),
        ),
    ]