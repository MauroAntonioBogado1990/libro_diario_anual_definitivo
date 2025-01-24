# -*- coding:utf-8 -*-

{
    'name': 'libro diario anual',
    'version': '16.0',
    'depends': [
        'base','stock','account'
    ],
    'author': 'Mauro Bogado, Exemax',
    'website': 'www.exemax.com.ar',
    'summary': 'Modulo de que agrega el libro diario anual con el wizard',
    'category': 'Extra Tools',
    'description': '''
    Modulo de que agrega el libro diario anual con el wizard.
    ''',
    'data': [
        'views/libro_diario_anual.xml',
        'wizard/wizard.xml',
        'security/ir.model.access.csv',
       
    ],
}