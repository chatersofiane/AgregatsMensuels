# Generated by Django 3.2.6 on 2021-08-19 23:13

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0023_auto_20210820_0000'),
    ]

    operations = [
        migrations.AlterField(
            model_name='checkbox',
            name='canva',
            field=models.OneToOneField(on_delete=django.db.models.deletion.CASCADE, primary_key=True, related_name='checkbox_canva', serialize=False, to='bilans.canva'),
        ),
    ]
