# -*- coding: utf-8 -*-
__version__ = "0.1.0"

from collections import OrderedDict

from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.generic import BooleanObject, IndirectObject, NameObject


def _getFields(obj, tree=None, retval=None, fileobj=None):
    """
    Extracts field data if this PDF contains interactive form fields.
    The *tree* and *retval* parameters are for recursive use.

    :param fileobj: A file object (usually a text file) to write
        a report to on all interactive form fields found.
    :return: A dictionary where each key is a field name, and each
        value is a :class:`Field<PyPDF2.generic.Field>` object. By
        default, the mapping name is used for keys.
    :rtype: dict, or ``None`` if form data could not be located.
    """
    fieldAttributes = {
        "/FT": "Field Type",
        "/Parent": "Parent",
        "/T": "Field Name",
        "/TU": "Alternate Field Name",
        "/TM": "Mapping Name",
        "/Ff": "Field Flags",
        "/V": "Value",
        "/DV": "Default Value",
    }
    if retval is None:
        retval = OrderedDict()
        catalog = obj.trailer["/Root"]
        # get the AcroForm tree
        if "/AcroForm" in catalog:
            tree = catalog["/AcroForm"]
        else:
            return None
    if tree is None:
        return retval

    obj._checkKids(tree, retval, fileobj)
    for attr in fieldAttributes:
        if attr in tree:
            # Tree is a field
            obj._buildField(tree, retval, fileobj, fieldAttributes)
            break

    if "/Fields" in tree:
        fields = tree["/Fields"]
        for f in fields:
            field = f.getObject()
            obj._buildField(field, retval, fileobj, fieldAttributes)

    return retval


def get_form_fields(infile):
    infile = PdfFileReader(open(infile, "rb"))
    fields = _getFields(infile)
    return OrderedDict((k, v.get("/V", "")) for k, v in fields.items())


def set_need_appearances_writer(writer: PdfFileWriter):
    # See 12.7.2 and 7.7.2 for more information: http://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/PDF32000_2008.pdf
    try:
        catalog = writer._root_object
        # get the AcroForm tree
        if "/AcroForm" not in catalog:
            writer._root_object.update(
                {
                    NameObject("/AcroForm"): IndirectObject(
                        len(writer._objects), 0, writer
                    )
                }
            )

        need_appearances = NameObject("/NeedAppearances")
        writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
        return writer

    except Exception as e:
        print("set_need_appearances_writer() catch : ", repr(e))
        return writer


def update_form_values(infile, outfile, newvals=None):
    """Update form fields
    infile: source PDF file to fill
    outfile: PDF file to save as
    newvals: dictionary to fill PDF = {'field_name': 'value'}
        when empty it will get all fields in source PDF and fill them with its own field names
    """
    pdf = PdfFileReader(open(infile, "rb"))
    writer = PdfFileWriter()

    # https://github.com/mstamy2/PyPDF2/issues/355#issuecomment-360569759
    # Fix for fields not showing up in Adobe Reader unless you click the field
    set_need_appearances_writer(writer)
    if "/AcroForm" in writer._root_object:
        writer._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )

    for i in range(pdf.getNumPages()):
        page = pdf.getPage(i)
        try:
            if newvals:
                writer.updatePageFormFieldValues(page, newvals)
            else:
                all_page_vals = {
                    k: f"#{i} {k}={v}"
                    for i, (k, v) in enumerate(get_form_fields(infile).items())
                }
                writer.updatePageFormFieldValues(
                    page,
                    all_page_vals,
                )
            writer.addPage(page)
        except Exception as e:
            print(repr(e))
            writer.addPage(page)

    with open(outfile, "wb") as out:
        writer.write(out)
