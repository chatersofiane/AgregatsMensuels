# Generated by Django 3.2.6 on 2021-08-12 13:29

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('bilans', '0006_auto_20210812_1427'),
    ]

    operations = [
        migrations.AlterField(
            model_name='bilan',
            name='canva',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='bilan_canva', to='bilans.canva'),
        ),
    ]