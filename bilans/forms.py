from django.forms.models import inlineformset_factory
from .models import Canva, Bilan, Autre, Tresorerie, Production, Checkbox
from django import forms
from django.forms import ModelForm

CheckboxsFormset= inlineformset_factory(Canva, Checkbox, fields=('validation',), can_delete=False,max_num=1,
                                        widgets = {
                                            'validation' : forms.CheckboxInput(attrs={'style': 'width:40px;height:40px;'}),
                                        },
                                       
                                            
                                      labels = {"validation" : ""},
                                      
                                                )


BilansFormset = inlineformset_factory(Canva, Bilan, fields=('agregat', 'SCF', 'mois1', 'mois2'), can_delete=False, max_num=26, min_num=26,
    
                                         widgets = {
                                            'agregat' : forms.TextInput(attrs={'class': 'col-sm-3'}),
                                            'SCF' : forms.TextInput(attrs={'class': 'col-sm-1'}),
                                            'mois1' : forms.NumberInput(attrs={'class': 'col-sm-1'}),
                                            'mois2' : forms.NumberInput(attrs={'class': 'col-sm-1'}),            
                                            
                                            
                                            },
                                            
                                            
                                      labels = {"agregat": "", "SCF": "", "mois1": "", "mois2": ""},
                                       
                                                )

    
AutresFormset = inlineformset_factory(Canva, Autre, fields=('autreag', 'SCF', 'finmoisn1', 'finmoisn','observation'), can_delete=False, max_num=14, min_num=14,  widgets = {
                                            'autreag' : forms.TextInput(attrs={'class': 'col-sm-2'}),
                                            'SCF' : forms.TextInput(attrs={'class': 'col-sm-1'}),
                                            'finmoisn1' : forms.NumberInput(attrs={'class': 'col-sm-2'}),
                                            'finmoisn' : forms.NumberInput(attrs={'class': 'col-sm-2'}),
                                            'observation' : forms.TextInput(attrs={'class': 'col-sm-4'}),
                                            }, labels = {"autreag": "", "SCF": "", "finmoisn1": "", "finmoisn": "", "ecart": "", "evolution": "", "observation": ""},
                                                )


TresoreriesFormset = inlineformset_factory(Canva, Tresorerie, fields=('SCF', 'banques', 'moism1', 'moism', 'observation'), can_delete=False, max_num=11, min_num=11, widgets = {
                                            'SCF' : forms.TextInput(attrs={'class': 'col-sm-1'}),
                                            'banques' : forms.TextInput(attrs={'class': 'col-sm-4'}),
                                            'moism1' : forms.NumberInput(attrs={'class': 'col-sm-1'}),
                                            'moism' : forms.NumberInput(attrs={'class': 'col-sm-1'}),
                                            'observation' : forms.TextInput(attrs={'class': 'col-sm-4'}),
                                            }, labels = {"SCF": "", "banques": "", "moism1": "", "moism": "", "observation": ""},
                                                )


ProductionsFormset = inlineformset_factory(Canva, Production, fields=('produit', 'unité', 'mois1', 'mois2'), can_delete=False, max_num=23, min_num=23, widgets = {
                                            'produit' : forms.TextInput(attrs={'class': 'col-sm-3'}),
                                            'unité' : forms.TextInput(attrs={'class': 'col-sm-1'}),
                                            'mois1' : forms.NumberInput(attrs={'class': 'col-sm-2'}),
                                            'mois2' : forms.NumberInput(attrs={'class': 'col-sm-2'}),   
                                            }, labels = {"produit": "", "unité": "", "mois1": "", "mois2": ""},
                                                )





