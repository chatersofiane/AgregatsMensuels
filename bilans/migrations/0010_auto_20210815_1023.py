# Generated by Django 3.2.6 on 2021-08-15 09:23

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0009_auto_20210815_0944'),
    ]

    operations = [
        migrations.AlterField(
            model_name='production',
            name='ecart1',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='ecart2',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='evolution1',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='evolution2',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='finmois1',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='finmois2',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='mois1',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='mois2',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='produit',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
        migrations.AlterField(
            model_name='production',
            name='unité',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
        migrations.AlterField(
            model_name='tresorerie',
            name='SCF',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
        migrations.AlterField(
            model_name='tresorerie',
            name='banques',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
        migrations.AlterField(
            model_name='tresorerie',
            name='moism',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='tresorerie',
            name='moism1',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='tresorerie',
            name='observation',
            field=models.CharField(blank=True, max_length=255, null=True),
        ),
    ]
