# test_excelAddinGenerator.py

import pytest
from excelAddinGenerator import *

def test_success_from_bin():
    # python excelAddinGenerator.py ./vbaProject.bin ./success_bin.xlam
    excelAddinGenerator.createFromBin("./vbaProject.bin", "success_bin.xlam")
    
def test_success_from_xlam():
    # python excelAddinGenerator.py ./test.xlam ./succes_xlam.xlam
    excelAddinGenerator.createFromBin("./test.xlam", "success_xlam.xlam")

#def test_bin_not_bin():
    # python excelAddinGenerator.py ./notbin.bin ./fail.xlam

#def test_xlam_not_zip():
    # python excelAddinGenerator.py ./notzip.xlam ./fail.xlam
