# Generated by Django 3.2.6 on 2021-09-05 07:46

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0063_auto_20210902_1320'),
    ]

    operations = [
        migrations.AlterField(
            model_name='bilan',
            name='evolution1',
            field=models.DecimalField(blank=True, decimal_places=2, max_digits=20, null=True),
        ),
        migrations.AlterField(
            model_name='bilan',
            name='evolution2',
            field=models.DecimalField(blank=True, decimal_places=5, max_digits=20, null=True),
        ),
    ]
