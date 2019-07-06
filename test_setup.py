import setup
import os
import pytest


def clean_up(dir, list_of_locs):
    for loc in os.listdir(dir):
        if loc in list_of_locs:
            os.removedirs(os.path.join(dir, loc))


def check_made(dir, list_of_locs):
    for loc in os.listdir(loc):
        if loc in list_of_locs:
            list_of_locs.remove(loc)
        else:
            print(loc)
    assert list_of_locs == []


def test_main():
    make_in_main = ["fotos.xlsx", "small"]
    loc = "/home/tako/testroom/excel_fotos/fotos"
    clean_up(loc, make_in_main)

    setup.main(pathname=loc, excelfile=os.path.join(loc, "fotos.xlsx"))
    check_made(loc, make_in_main)
