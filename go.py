# -*- coding: utf-8 -*-

import locale
import os
import re
import warnings
from pprint import pprint

from dotenv import load_dotenv
from openpyxl import load_workbook

from accounts import update_form_values

if __name__ == "__main__":

    load_dotenv()

    locale.setlocale(locale.LC_ALL, os.getenv('locale', 'en'))
    in_folder = os.getenv("source_folder")
    out_folder = os.getenv("out_folder")
    pdf_file_name = "S-26-S.pdf"

    infile = f"{in_folder}/{pdf_file_name}"

    xlsx_file_name = f"{in_folder}/S-26-S.xlsx"

    # suppress the user warnings when opening the xlsx file
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=UserWarning)
        wb = load_workbook(xlsx_file_name, data_only=True)
    # ws = wb.active

    for ws_name in wb.sheetnames:
        outfile = f"{out_folder}/{ws_name}-{pdf_file_name}"

        ws = wb[ws_name]

        # tbl_ref = ws.tables['Table2'].ref
        # tbl_data = ws[tbl_ref]
        row_start = 3
        row_end = 54

        year = ws.cell(row=1, column=24).value
        month = ws.cell(row=1, column=23).value
        month_end = ws.cell(row=1, column=22).value

        # get totals
        re_anterior = ws.cell(row=2, column=21).value  # recibido saldo anterior
        cp_anterior = ws.cell(row=3, column=21).value  # cuenta principal saldo anterior
        oo_anterior = ws.cell(row=4, column=21).value  # "otra" saldo anterior

        re_ent_total = ws.cell(row=55, column=5).value
        re_sal_total = ws.cell(row=55, column=8).value
        re_sal_final = float(re_anterior or 0) + float(re_ent_total or 0) - float(re_sal_total or 0)
        cp_ent_total = ws.cell(row=55, column=10).value
        cp_sal_total = ws.cell(row=55, column=12).value
        cp_sal_final = float(cp_anterior or 0) + float(cp_ent_total or 0) - float(cp_sal_total or 0)
        oo_ent_total = ws.cell(row=55, column=14).value
        oo_sal_total = ws.cell(row=55, column=16).value
        oo_sal_final = float(oo_anterior or 0) + float(oo_ent_total or 0) - float(oo_sal_total or 0)

        total = re_sal_final + cp_sal_final + oo_sal_final

        # format the totals in locale
        re_ent_total_fmt = locale.format_string("%.2f", re_ent_total) if re_ent_total else ""
        re_sal_total_fmt = locale.format_string("%.2f", re_sal_total) if re_sal_total else ""
        cp_ent_total_fmt = locale.format_string("%.2f", cp_ent_total) if cp_ent_total else ""
        cp_sal_total_fmt = locale.format_string("%.2f", cp_sal_total) if cp_sal_total else ""
        oo_ent_total_fmt = locale.format_string("%.2f", oo_ent_total) if oo_ent_total else ""
        oo_sal_total_fmt = locale.format_string("%.2f", oo_sal_total) if oo_sal_total else ""

        re_anterior_fmt = locale.format_string("%.2f", re_anterior) if re_anterior else ""
        cp_anterior_fmt = locale.format_string("%.2f", cp_anterior) if cp_anterior else ""
        oo_anterior_fmt = locale.format_string("%.2f", oo_anterior) if oo_anterior else ""

        re_sal_final_fmt = locale.format_string("%.2f", re_sal_final) if re_sal_final else ""
        cp_sal_final_fmt = locale.format_string("%.2f", cp_sal_final) if cp_sal_final else ""
        oo_sal_final_fmt = locale.format_string("%.2f", oo_sal_final) if oo_sal_final else ""

        total_fmt = locale.format_string("%.2f", total) if total else ""

        # prepare the data_dict with initial values
        data_dict = {
            "900_1_Text_C": os.getenv('cong'),
            "900_2_Text_C": os.getenv('city'),
            "900_3_Text_C": os.getenv('state'),
            "900_4_Text_C": month,
            "900_5_Text_C": year,
            "901_53_S26TotalValue": re_ent_total_fmt,
            "901_106_S26TotalValue": re_sal_total_fmt,
            "902_53_S26TotalValue": cp_ent_total_fmt,
            "902_106_S26TotalValue": cp_sal_total_fmt,
            "903_53_S26TotalValue": oo_ent_total_fmt,
            "903_106_S26TotalValue": oo_sal_total_fmt,
            "904_28_Text_C": month_end,  # month end
            "904_29_S26Amount": re_anterior_fmt,  # re_saldo_anterior FILL IN
            "904_30_S26TotalAmount": re_ent_total_fmt,  # re_entrada
            "904_31_S26TotalAmount": re_sal_total_fmt,  # re_salida
            "904_32_S26TotalAmount": re_sal_final_fmt,  # re_salida
            "904_33_S26Amount": cp_anterior_fmt,  # cp_saldo_anterior FILL IN
            "904_34_S26TotalAmount": cp_ent_total_fmt,  # cp_entrada
            "904_35_S26TotalAmount": cp_sal_total_fmt,  # cp_salida
            "904_36_S26TotalAmount": cp_sal_final_fmt,  # cp_salida
            "904_38_S26Amount": oo_anterior_fmt,  # oo_saldo_anterior FILL IN
            "904_39_S26TotalAmount": oo_ent_total_fmt,  # oo_entrada
            "904_40_S26TotalAmount": oo_sal_total_fmt,  # oo_salida
            "904_41_S26TotalAmount": oo_sal_final_fmt,  # oo_salida
            "904_42_S26TotalAmount": total_fmt,
        }

        # columns to add to the data_dict
        cols = [
            {"name": "fecha", "field": "900_7_Text_C", "col_num": 1, "text": True},
            {"name": "descr", "field": "900_59_Text", "col_num": 2, "text": True},
            {"name": "ct", "field": "900_111_Text_C", "col_num": 4, "text": True},
            {"name": "r_ent", "field": "901_1_S26Value", "col_num": 5, "text": False},
            {"name": "r_sal", "field": "901_54_S26Value", "col_num": 8, "text": False},
            {"name": "cp_ent", "field": "902_1_S26Value", "col_num": 10, "text": False},
            {"name": "cp_sal", "field": "902_54_S26Value", "col_num": 12, "text": False},
            {"name": "o_ent", "field": "903_1_S26Value", "col_num": 14, "text": False},
            {"name": "o_sal", "field": "903_54_S26Value", "col_num": 16, "text": False},
        ]

        # add the columns to the data_dict
        for col in cols:
            cell_data = ws.cell(row=row_start, column=col["col_num"])
            field = col["field"]
            for n in range(row_start, row_end):
                cell_data = ws.cell(row=n, column=col["col_num"])
                if cell_data.value:
                    data_dict[field] = (
                        cell_data.value
                        if col["text"]
                        else locale.format_string("%.2f", cell_data.value)
                    )  # col + "_" + str(n)
                del cell_data
                field = re.sub(
                    r"_([^_][0-9]*)_",
                    lambda x: f"_{str(int(x.group(1))+1).zfill(len(x.group(1)))}_",
                    field,
                )
        # pprint(data_dict)
        update_form_values(
            infile,
            outfile,
            data_dict,
            # {"my_fieldname_1": "My Value", "my_fieldname_2": "My Another value"},
        )
