# test_excelAddinGenerator.py

import pytest
from excelAddinGenerator.main import *

def test_success_from_bin():
    # python excelAddinGenerator.py ./vbaProject.bin ./success_bin.xlam
    createFromBin("tests/vbaProject.bin", "success_bin.xlam")
    
def test_fail_from_bin():
    with pytest.raises(Exception) as e_info:
    # python excelAddinGenerator.py ./test.xlam ./succes_xlam.xlam
    createFromBin("tests/test.xlam", "success_xlam.xlam")

#def test_bin_not_bin():
    # python excelAddinGenerator.py ./notbin.bin ./fail.xlam

#def test_xlam_not_zip():
    # python excelAddinGenerator.py ./notzip.xlam ./fail.xlam
