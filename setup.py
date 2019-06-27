# https://stackoverflow.com/a/25817258
import numpy as np

# import bpy
# from bpy_extras.io_utils import ImportHelper, ExportHelper
import tkinter
from tkinter.filedialog import askdirectory
from xlsxwriter.workbook import Workbook
from PIL import Image
from resizeimage import resizeimage
from PIL.ExifTags import TAGS


def add_lookup_table(workbook: Workbook):
    worksheet = workbook.add_worksheet("table")
    with open("lookup_table.txt", "r") as f:
        for i, line in enumerate(f):
            e, b, h, u = line.split("	")
            worksheet.write("A" + str(i + 1), e)
            worksheet.write("B" + str(i + 1), b)
            worksheet.write("C" + str(i + 1), h)
            worksheet.write("D" + str(i + 1), u)


def insert_image(worksheet, image_path_small, image_path_list, j):
    worksheet.insert_image(
        "F" + str(j),
        image_path_small,
        {
            "positioning": 1,
            "x_scale": 0.4,
            "y_scale": 0.4,
            "y_offset": 2,
            "x_offset": 2,
        },
    )
    worksheet.write_comment("F" + str(j), image_path_list[-1])


if __name__ == "__main__":
    # register()
    import datetime
    from tkinter.filedialog import askdirectory, asksaveasfilename
    import os

    pathname = askdirectory()
    images = os.listdir(pathname)
    print(pathname)
    list_name = pathname.split(os.sep)
    name = list_name[-1]
    print(name)
    try:
        os.mkdir(os.path.join(pathname, "small"))
    except:
        print("path already exicist")
    exelfile = asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel file", ".xlsx")],
        initialdir=pathname,
        initialfile=name,
    )
    # print(os.listdir(pathname))
    # Create an new Excel file and add a worksheet.
    workbook = Workbook(exelfile)
    worksheet = workbook.add_worksheet()

    format_wrap = workbook.add_format()
    format_wrap.set_text_wrap()

    # Widen the first column to make the text clearer.
    worksheet.set_column("A:A", None, None, {"hidden": 1})
    worksheet.set_column("F:F", 30)
    worksheet.set_column("B:B", 12)
    worksheet.set_column("H:H", 25)
    worksheet.set_column("L:L", 25, format_wrap)
    worksheet.set_column("G:G", 25, format_wrap)

    # TODO worksheet.autofilter() autofilter(first_row, first_col, last_row, last_col)
    # TODO conditional formating
    # TODO headers
    # TODO dropdown
    # TODO ??info???

    worksheet.write("A1", "Key")
    worksheet.write("B1", "Datum")
    worksheet.write("C1", "Tijd")
    worksheet.write("D1", "Afdeling")
    worksheet.write("E1", "Locatie")
    worksheet.write("F1", "Foto")
    worksheet.write("G1", "Risicobeschrijving")
    worksheet.write("H1", "Risicogebied")
    with open("risicoscore.txt", "r") as f:
        line = f.readline()
        for col, head in zip(["I", "J", "K", "L", "M", "N", "O"], line.split("|")):
            worksheet.write(col + "1", head)

    # format910 = workbook.add_format({"bg_color": "#DF0101", "font_color": "#000000"})

    # format78 = workbook.add_format({"bg_color": "#FF8000", "font_color": "#000000"})

    # format56 = workbook.add_format({"bg_color": "#FFFF00", "font_color": "#000000"})

    # format34 = workbook.add_format({"bg_color": "#80FF00", "font_color": "#000000"})

    # format2 = workbook.add_format({"bg_color": "#088A08", "font_color": "#000000"})

    j = 2
    date_format = workbook.add_format({"num_format": "dd/mm/yyyy"})
    time_format = workbook.add_format({"num_format": "hh:mm"})

    # Write another conditional format over the same range.
    worksheet.conditional_format(
        "K2:K1000", {"type": "cell", "criteria": "=", "value": 2, "format": format2}
    )
    # worksheet.conditional_format(
    #     "K2:K1000",
    #     {
    #         "type": "cell",
    #         "criteria": "between",
    #         "minimum": 2.1,
    #         "maximum": 4.5,
    #         "format": format34,
    #     },
    # )
    # worksheet.conditional_format(
    #     "K2:K1000",
    #     {
    #         "type": "cell",
    #         "criteria": "between",
    #         "minimum": 4.6,
    #         "maximum": 6.5,
    #         "format": format56,
    #     },
    # )
    # worksheet.conditional_format(
    #     "K2:K1000",
    #     {
    #         "type": "cell",
    #         "criteria": "between",
    #         "minimum": 6.6,
    #         "maximum": 8.5,
    #         "format": format78,
    #     },
    # )
    # worksheet.conditional_format(
    #     "K2:K1000",
    #     {
    #         "type": "cell",
    #         "criteria": "between",
    #         "minimum": 8.6,
    #         "maximum": 11,
    #         "format": format910,
    #     },
    # )

    for i, image in enumerate(images):

        if image.split(".")[-1] in ["jpg", "JPG"]:
            image_path = os.path.join(pathname, image)
            image_path_list = image_path.split(os.sep)
            image_path_list.insert(-1, "small")
            image_path_small = os.sep.join(image_path_list)
            im = Image.open(image_path)
            im.thumbnail((512, 512), Image.ANTIALIAS)
            im.save(image_path_small)

            worksheet.set_row(j - 1, 170)
            try:
                tags = Image.open(image_path)._getexif()
                # print(tags)
                date = tags[36867]

            except:
                date = "01-01-2010 00:00:00"
            ret = {}
            for tag, value in tags.items():
                decoded = TAGS.get(tag, tag)
                ret[decoded] = value
            print(ret)
            worksheet.write("A" + str(j), str(date.replace(" ", "")))
            date_ = datetime.datetime.strptime(date.split(" ")[0], "%Y:%m:%d")
            worksheet.write_datetime("B" + str(j), date_, date_format)

            time_ = datetime.datetime.strptime(date.split(" ")[1], "%H:%M:%S")
            worksheet.write_datetime("C" + str(j), time_, time_format)
            insert_image(worksheet, image_path_small, image_path_list, j)

            worksheet.data_validation(
                "H" + str(j),
                {
                    "validate": "list",
                    "source": [
                        "arbeidsplaatsen",
                        "gevaarlijke stoffen",
                        "fysieke belasting",
                        "fysische omstandigheden",
                        "arbeidsmiddelen",
                        "PBM en VG signalering",
                    ],
                },
            )

            worksheet.data_validation(
                "I" + str(j), {"validate": "list", "source": [1, 2, 3, 4, 5]}
            )

            worksheet.data_validation(
                "J" + str(j), {"validate": "list", "source": [1, 2, 3, 4, 5]}
            )
            # "=ZOEKEN(B33;table!G36:G91;table!H36:H91)"
            worksheet.write_formula("K" + str(j), "=I" + str(j) + "+J" + str(j))

            j += 1
    worksheet.autofilter(0, 0, 1000, 11)
    worksheet.freeze_panes(1, 0)
    add_lookup_table(workbook)
    workbook.close()
