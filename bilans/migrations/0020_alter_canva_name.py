# Generated by Django 3.2.6 on 2021-08-19 21:48

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0019_bilan_validation'),
    ]

    operations = [
        migrations.AlterField(
            model_name='canva',
            name='name',
            field=models.CharField(blank=True, max_length=255, null=True, unique=True),
        ),
    ]
