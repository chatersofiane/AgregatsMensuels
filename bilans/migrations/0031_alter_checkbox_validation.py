# Generated by Django 3.2.6 on 2021-08-20 00:21

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0030_auto_20210820_0106'),
    ]

    operations = [
        migrations.AlterField(
            model_name='checkbox',
            name='validation',
            field=models.BooleanField(default=False),
        ),
    ]
