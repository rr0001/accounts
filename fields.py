# -*- coding: utf-8 -*-

import locale
import os
from pprint import pprint

from dotenv import load_dotenv

from accounts import update_form_values

if __name__ == "__main__":

    load_dotenv()

    locale.setlocale(locale.LC_ALL, os.getenv("locale", "en"))
    in_folder = "sources"
    out_folder = "out"
    pdf_file_name = "S-26-S.pdf"

    infile = f"{in_folder}/{pdf_file_name}"
    outfile = f"{out_folder}/fields_{pdf_file_name}"

    # # stdout the fields
    # fields = get_form_fields(pdf_file_name)
    # pprint(fields)

    # enumerate & fill the fields with their own names
    update_form_values(infile, outfile)
