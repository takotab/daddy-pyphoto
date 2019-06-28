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

sep = "|"


def add_lookup_table(workbook: Workbook):
    worksheet = workbook.add_worksheet("table")
    with open("lookup_table.txt", "r") as f:
        for i, line in enumerate(f):
            e, b, h, u = line.split("	")
            if i == 0:
                worksheet.write("A" + str(i + 1), e)
                worksheet.write("B" + str(i + 1), b)
                worksheet.write("C" + str(i + 1), h)
                worksheet.write("D" + str(i + 1), u)
            else:
                worksheet.write_number("A" + str(i + 1), int(e))
                worksheet.write_number("B" + str(i + 1), int(b))
                worksheet.write_number("C" + str(i + 1), int(h))
                worksheet.write_number("D" + str(i + 1), int(u))


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


class Column(object):
    def __init__(self, column: int):
        self.col_letter = ["I", "J", "K", "L", "M", "N", "O"][column]
        with open("risicoscore.txt", "r") as f:
            self.head = (
                f.readline().split(sep)[column].replace("\n", "").replace(" ", "")
            )
            self.comment = ""
            self.valid = []

            if self.head in ["hash", "risico"]:
                return
            else:
                for row_num, line in enumerate(f):
                    r = line.split(sep)[column]
                    if not r == "":
                        self.comment += str(row_num) + ")" + r + "  "
                        self.valid.append(row_num)

    def formula(self, x: int):
        if self.head == "hash":
            return (
                "=100*I"
                + str(x)
                + "+J"
                + str(x)
                + "+K"
                + str(x)
                + "+L"
                + str(x)
                + "+M"
                + str(x)
            )

        elif self.head == "risico":
            return "=ZOEKEN(B" + str(x) + ",table!C2:C200,table!D2:D200)"


def make_cols():
    return [Column(i) for i in range(7)]


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
    worksheet.set_column("J:J", 25)
    worksheet.set_column("K:K", 25)
    worksheet.set_column("M:M", 25)
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

    j = 2
    date_format = workbook.add_format({"num_format": "dd/mm/yyyy"})
    time_format = workbook.add_format({"num_format": "hh:mm"})
    cols = make_cols()

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

            for col in cols:
                print(col.head, col.comment, col.col_letter)
                if col.head in ["hash", "risico"]:
                    print(col.head, col.formula(j))
                    worksheet.write_formula(col.col_letter + str(j), col.formula(j))
                else:
                    worksheet.data_validation(
                        col.col_letter + str(j),
                        {"validate": "list", "source": col.valid},
                    )
                    worksheet.write_comment(col.col_letter + str(j), col.comment)

                # else:
            j += 1
    worksheet.autofilter(0, 0, 1000, 11)
    worksheet.freeze_panes(1, 0)
    add_lookup_table(workbook)
    workbook.close()
