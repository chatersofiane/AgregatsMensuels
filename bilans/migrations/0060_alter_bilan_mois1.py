# Generated by Django 3.2.6 on 2021-09-02 09:39

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0059_alter_bilan_mois1'),
    ]

    operations = [
        migrations.AlterField(
            model_name='bilan',
            name='mois1',
            field=models.DecimalField(blank=True, decimal_places=5, default=0, max_digits=20, null=True),
        ),
    ]
