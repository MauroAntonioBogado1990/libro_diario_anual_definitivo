# -*- coding: utf-8 -*-
from odoo import models, fields, api
import base64
import io
import xlsxwriter
from collections import defaultdict

import logging # Importar el módulo logging
_logger = logging.getLogger(__name__)

class AddBookAnualWizard(models.TransientModel):
    _name = "wizard.book"

    date_start = fields.Date(string='Fecha Inicio')
    date_end = fields.Date(string='Fecha Fin')
    journal_ids = fields.Many2many('account.journal', string='Diarios')
    number_journal = fields.Char(string="Libro Diario N°")

    def action_confirm(self):
        for journal in self.journal_ids:
            _logger.info(f"Diario seleccionado: {journal.name}")

        # Llamada a la función generate_xlsx_report y obtener el archivo
        file_data, file_name = self.generate_xlsx_report()

        # Crear un adjunto con el archivo generado
        attachment = self.env['ir.attachment'].create({
            'name': file_name,
            'type': 'binary',
            'datas': base64.b64encode(file_data),
            'store_fname': file_name,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })

        # Devolver acción para descargar el archivo
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%s?download=true' % attachment.id,
            'target': 'self',
        }
    
    def generate_xlsx_report(self):
        start_date = self.date_start
        end_date = self.date_end
        journal_ids = self.journal_ids
        number_journal = self.number_journal

        moves = self.env['account.move'].search([
            ('state', '=', 'posted'),
            ('date', '>=', start_date),
            ('date', '<=', end_date),
            ('journal_id', 'in', journal_ids.ids)
        ],
        order='date asc, journal_id asc, name asc, id asc'
        )

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        data = {}  # Asumiendo que "data" contiene información adicional que necesitas pasar
        # Incluir el campo number_journal en el diccionario data
        data = {'number_journal': number_journal, 'date_start': str(start_date)}
        self.env['report.accounting_report.accounting_report_busch'].generate_xlsx_report(workbook, data, moves)
        workbook.close()

        file_data = output.getvalue()
        output.close()

        file_name = 'Libro_Diario_Anual.xlsx'

        return file_data, file_name
    

class BookDaily(models.AbstractModel):
    _name = 'report.accounting_report.accounting_report_busch'
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, data, moves):
        report_name = 'Reporte Contable'

        year = fields.Date.from_string(data['date_start']).year
        ejercicio = year - 1983  # 1984 fue el Ejercicio 1


        sheet = workbook.add_worksheet(report_name[:31])
        h = "#"
        money_format = workbook.add_format({'num_format': "$ 0" + h + h + '.' + h + h + ',' + h + h})
        bold = workbook.add_format({'bold': True})
        sheet.write(0, 2, '', bold)
        sheet.write(1, 2, '')
        sheet.write(2, 2, '', bold)
        #sheet.write(2, 3, data['number_journal'])
        sheet.write(3, 0, 'Ejercicio')
        sheet.write(3, 1, ejercicio)
        sheet.write(4, 0, 'Diario General', bold)

        sheet.write(5, 0, 'Número')
        sheet.write(5, 2, 'Fecha')
        sheet.write(5, 3, 'Concepto')

        centrado = workbook.add_format()
        centrado.set_align('vcenter')
        border_box = workbook.add_format({'border': 1})
#        sheet.write(7, 0, 'DEBE', centrado)
#        sheet.write(7, 3, 'HABER', centrado)

        debe_header_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bold': True
        })
        haber_header_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'bold': True
        })
        sheet.merge_range(6, 0, 6, 2, 'DEBE', debe_header_format)
        sheet.merge_range(6, 3, 6, 5, 'HABER', haber_header_format)

        border_bottom_format = workbook.add_format({
            'bottom': 1,  # Línea más gruesa
            'bold': True,
            'align': 'center',
            'valign': 'vcenter'
        })

        sheet.write(8, 0, 'Cuenta', border_bottom_format)
        sheet.write(8, 1, '', border_bottom_format)  # Espacio vacío con borde
        sheet.write(8, 2, 'Importe', border_bottom_format)
        sheet.write(8, 3, 'Cuenta', border_bottom_format)
        sheet.write(8, 4, '', border_bottom_format)  # Espacio vacío con borde
        sheet.write(8, 5, 'Importe', border_bottom_format)

        debe_index = 10
        haber_index = 10

        total_debe = 0
        total_haber = 0

        previous_values = {"journal_id": None, "date": None, "ref": None}
        debe_accounts = defaultdict(lambda: {'name': '', 'amount': 0.0})
        haber_accounts = defaultdict(lambda: {'name': '', 'amount': 0.0})

        for move in moves:
            current_row = max(debe_index, haber_index)


            # Verificar si journal_id, date o ref han cambiado
            if (move.journal_id.code != previous_values["journal_id"] or
                move.date.strftime("%Y-%m-%d") != previous_values["date"] or
                (move.ref or move.name) != previous_values["ref"]):

                # Imprimir los valores acumulados de la entrada anterior
                if previous_values["journal_id"] is not None:
                    for code, data in debe_accounts.items():
                        sheet.write(debe_index, 0, code)
                        sheet.write(debe_index, 1, data['name'])
                        sheet.write(debe_index, 2, data['amount'], money_format)
                        debe_index += 1

                    for code, data in haber_accounts.items():
                        sheet.write(haber_index, 3, code)
                        sheet.write(haber_index, 4, data['name'])
                        sheet.write(haber_index, 5, data['amount'], money_format)
                        haber_index += 1

                    max_index = max(debe_index, haber_index) + 1
                    debe_index = max_index

                    current_row = max(debe_index, haber_index)  # Ajustar la fila actual después de imprimir los detall>
                # Escribir nueva línea de encabezado de entrada
                #sheet.write(current_row, 0, move.journal_id.code)
                #sheet.write(current_row, 9, move.date.strftime("%Y-%m-%d"))
                #if move.date.strftime("%Y-%m-%d") != previous_values["date"]:
                border_right_format = workbook.add_format({
                    'top': 1,
                    'bottom': 1,
                    'right': 1,
                    'bold': False,
                    'align': 'left',
                    'valign': 'vcenter'
                })
                border_inner_format = workbook.add_format({
                    'top': 1, 'bottom': 1,  # Bordes solo arriba y abajo
                    'bold': False,
                    'align': 'center',
                    'valign': 'vcenter'
                })

                border_left_format = workbook.add_format({
                    'left': 1,  # Borde solo en el lado izquierdo
                    'top': 1,  # Borde superior
                    'bottom': 1,  # Borde inferior
                    'bold': False,
                    'align': 'center',
                    'valign': 'vcenter'
                })

                sheet.write(current_row, 0, move.name, border_left_format)  # Borde en todos los lados
                sheet.write(current_row, 1, '', border_inner_format)  # Sin bordes laterales
                sheet.write(current_row, 2, move.date.strftime("%Y-%m-%d"), border_inner_format)  # Fecha sin bordes la>#                sheet.write(current_row, 3, move.ref or move.name, border_inner_format)  # Concepto sin bordes lateral>#                sheet.write(current_row, 4, '', border_inner_format)  # Sin bordes laterales                           #                sheet.write(current_row, 5, '', border_outer_format)  # Borde en todos los lados
                sheet.merge_range(current_row, 3, current_row, 5, '', border_right_format)
#                sheet.write(current_row, 0, move.journal_id.code)
#                sheet.write(current_row, 2, move.date.strftime("%Y-%m-%d"))  # Fecha en negrita
#                sheet.write(current_row, 3, move.ref or move.name)
                current_row += 1
                debe_index = current_row
                haber_index = current_row

                #sheet.write(current_row, 3, move.ref or move.name)

                # Actualizar valores previos
                previous_values["journal_id"] = move.journal_id.code
                previous_values["date"] = move.date.strftime("%Y-%m-%d")

                previous_values["ref"] = move.ref or move.name

                debe_accounts = defaultdict(lambda: {'name': '', 'amount': 0.0})
                haber_accounts = defaultdict(lambda: {'name': '', 'amount': 0.0})
            
            for line in move.line_ids:
                if line.debit > 0:
                    debe_accounts[line.account_id.code]['name'] = line.account_id.name
                    debe_accounts[line.account_id.code]['amount'] += line.debit
                    total_debe += line.debit
                elif line.credit > 0:
                    haber_accounts[line.account_id.code]['name'] = line.account_id.name
                    haber_accounts[line.account_id.code]['amount'] += line.credit
                    total_haber += line.credit

            if move == moves[-1]:
                for code, data in debe_accounts.items():
                    sheet.write(debe_index, 0, code)
                    sheet.write(debe_index, 1, data['name'])
                    sheet.write(debe_index, 2, data['amount'], money_format)
                    debe_index += 1

                for code, data in haber_accounts.items():
                    sheet.write(haber_index, 3, code)
                    sheet.write(haber_index, 4, data['name'])
                    sheet.write(haber_index, 5, data['amount'], money_format)
                    haber_index += 1
        # Escribir totales
        final_row = max(debe_index, haber_index) + 1
        sheet.write(final_row, 2, total_debe, money_format)
        sheet.write(final_row, 5, total_haber, money_format)
        sheet.write(9, 2, total_debe, money_format)
        sheet.write(9, 5, total_haber, money_format)

        # Configurar la hoja para imprimir en tamaño A4
        sheet.set_paper(9)  # 9 corresponde a A4
        sheet.fit_to_pages(1, 0)  # Ajustar a una página de ancho

        # Ajustar los márgenes y la orientación
        sheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
        sheet.set_portrait()  # Orientación vertical

        # Ajustar ancho de columnas
        sheet.set_column('A:D', 20)
        sheet.set_column('E:H', 20)
        sheet.set_column('I:L', 20)            
