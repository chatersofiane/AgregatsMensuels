# Generated by Django 3.2.6 on 2021-08-18 10:17

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0018_delete_validation'),
    ]

    operations = [
        migrations.AddField(
            model_name='bilan',
            name='validation',
            field=models.BooleanField(blank=True, default=False, null=True),
        ),
    ]
