from django.db import migrations, models

class Migration(migrations.Migration):
    dependencies = [
        ('saisie', '0002_saisiemonthly_all_fields'),
    ]
    operations = [
        migrations.AddField(
            model_name='saisiemonthly',
            name='bonus_malus_saisi',
            field=models.DecimalField(
                max_digits=12, decimal_places=2, default=0,
                verbose_name="Bonus Malus (saisi manuellement)",
                help_text="Positif=excédent, négatif=déficit/prêt. Si 0: calculé auto."
            ),
        ),
    ]