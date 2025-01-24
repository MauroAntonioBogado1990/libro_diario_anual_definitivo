# -*- coding: utf-8 -*-
from odoo import models, api, fields
import json

class AddBookAnual(models.Model):
    
    _inherit = "account.move"

    