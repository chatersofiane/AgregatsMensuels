# Generated by Django 3.2.6 on 2021-08-11 08:54

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0002_alter_bilan_agregat'),
    ]

    operations = [
        migrations.AlterField(
            model_name='bilan',
            name='agregat',
            field=models.CharField(max_length=255),
        ),
    ]
