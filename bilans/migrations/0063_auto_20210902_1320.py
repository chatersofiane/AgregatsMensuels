# Generated by Django 3.2.6 on 2021-09-02 12:20

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0062_alter_bilan_evolution1'),
    ]

    operations = [
        migrations.AlterField(
            model_name='bilan',
            name='ecart1',
            field=models.IntegerField(blank=True, null=True),
        ),
        migrations.AlterField(
            model_name='bilan',
            name='ecart2',
            field=models.IntegerField(blank=True, null=True),
        ),
        migrations.AlterField(
            model_name='bilan',
            name='finmois1',
            field=models.IntegerField(blank=True, null=True),
        ),
        migrations.AlterField(
            model_name='bilan',
            name='finmois2',
            field=models.IntegerField(blank=True, null=True),
        ),
        migrations.AlterField(
            model_name='bilan',
            name='mois1',
            field=models.IntegerField(blank=True, null=True),
        ),
        migrations.AlterField(
            model_name='bilan',
            name='mois2',
            field=models.IntegerField(blank=True, null=True),
        ),
    ]
