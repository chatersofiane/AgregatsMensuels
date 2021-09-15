from django.db import models
from django.urls import reverse
from django.utils import timezone

# Create your models here.






class Canva(models.Model):
    name = models.CharField(null=True, blank=True, max_length=255)
    site = models.CharField(null=True, blank=True, max_length=255)
    mois = models.CharField(null=True, blank=True, max_length=255)
    année = models.IntegerField(null=True, blank=True)
    created = models.DateTimeField(null=True, blank=True, default=timezone.now)
    
    
    
    
    def get_absolute_url(self):
        return reverse('bilans:canva_detail', kwargs={'pk': self.pk})
    
    def __str__(self):
        return self.name
    
class Checkbox(models.Model):
    validation = models.BooleanField(default=True)
    canva = models.ForeignKey('Canva', null=True, blank=True, on_delete=models.CASCADE, related_name='checkbox_canva')
    
   
       


class Bilan(models.Model):
    agregat = models.CharField(null=True, blank=True, max_length=255)
    SCF = models.CharField(null=True, blank=True, max_length=255)
    mois1 = models.IntegerField(null=True, blank=True)
    mois2 =  models.IntegerField(null=True, blank=True)
    ecart1 =  models.IntegerField(null=True, blank=True)
    evolution1 = models.DecimalField(max_digits=20, decimal_places=2, null=True, blank=True)
    finmois1 =  models.IntegerField(null=True, blank=True)
    finmois2 =  models.IntegerField(null=True, blank=True)
    ecart2 =  models.IntegerField(null=True, blank=True)
    evolution2 = models.DecimalField(max_digits=20, decimal_places=2, null=True, blank=True)
    canva = models.ForeignKey('Canva', null=True, blank=True, on_delete=models.CASCADE, related_name='bilan_canva')
    
    
class Autre(models.Model):
    autreag = models.CharField(null=True, blank=True, max_length=255)
    SCF = models.CharField(null=True, blank=True, max_length=255)
    finmoisn1 = models.IntegerField(null=True, blank=True)
    finmoisn = models.IntegerField(null=True, blank=True)
    ecart = models.IntegerField(null=True, blank=True)
    evolution = models.DecimalField(max_digits=20, decimal_places=2, null=True, blank=True)
    observation = models.CharField(null=True, blank=True, max_length=255)
    canva = models.ForeignKey('Canva', null=False, blank=False, on_delete=models.CASCADE, related_name='autre_canva')
    
   
class Tresorerie(models.Model):
    SCF = models.CharField(null=True, blank=True, max_length=255)
    banques = models.CharField(null=True, blank=True, max_length=255)
    moism1 = models.IntegerField(null=True, blank=True)
    moism = models.IntegerField(null=True, blank=True)
    observation = models.CharField(null=True, blank=True, max_length=255)
    canva = models.ForeignKey('Canva', null=False, blank=False, on_delete=models.CASCADE, related_name='tresorerie_canva')    
    
    
    
class Production(models.Model):
    produit = models.CharField(null=True, blank=True, max_length=255)
    unité = models.CharField(null=True, blank=True, max_length=255)
    mois1 = models.IntegerField(null=True, blank=True)
    mois2 = models.IntegerField(null=True, blank=True)
    ecart1 = models.IntegerField(null=True, blank=True)
    evolution1 = models.DecimalField(max_digits=20, decimal_places=2, null=True, blank=True)
    finmois1 = models.IntegerField(null=True, blank=True)
    finmois2 = models.IntegerField(null=True, blank=True)
    ecart2 = models.IntegerField(null=True, blank=True)
    evolution2 = models.DecimalField(max_digits=20, decimal_places=2, null=True, blank=True)
    canva = models.ForeignKey('Canva', null=False, blank=False, on_delete=models.CASCADE, related_name='production_canva')
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    