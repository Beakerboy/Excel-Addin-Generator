# test_excelAddinGenerator.py

import pytest
from excelAddinGenerator.main import *

def test_success_from_bin():
    """Test that xlam is successfully generated from a OLE file"""
    createFromBin("tests/vbaProject.bin", "src/data", "success_bin.xlam")
    # Assert that xlam file is created
    
def test_fail_from_bin():
    """ Test that an exception is thrown if the file is not an OLE file"""
    with pytest.raises(Exception) as e_info:
        createFromBin("tests/test.xlam", "src/data", "success_xlam.xlam")

#def test_bin_not_bin():
    # python excelAddinGenerator.py ./notbin.bin ./fail.xlam

#def test_xlam_not_zip():
    # python excelAddinGenerator.py ./notzip.xlam ./fail.xlam
