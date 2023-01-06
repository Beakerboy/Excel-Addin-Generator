# test_excelAddinGenerator.py

import pytest
from excelAddinGenerator.main import *
from os.path import exists

def test_main():
    main(["./excelAddinGenerator/", "tests/vbaProject.bin", "success_bin.xlam"])

def test_success_from_bin():
    """Test that xlam is successfully generated from a OLE file"""
    createFromBin("tests/vbaProject.bin", "src/data", "success_bin.xlam")
    # Assert that xlam file is created
    assert exists("success_bin.xlam")
    #ToDo: assert that bin file within success_bin.xlam matches tests/vbaProject.bin
    
    createFromZip("success_bin.xlam", "src/data", "success_xlam.xlam")
    assert exists("success_xlam.xlam")
    #ToDo: assert that bin file within success_xlam.xlam matches bin file within success_bin.xlam
    
def test_not_bin_exception():
    """ Test that an exception is thrown if the bin file is not an OLE file"""
    with pytest.raises(Exception) as e_info:
        createFromBin("tests/blank.bin", "src/data", "./fail.xlam")
        
def test_xlam_not_zip():
    """ Test that an exception is thrown if the zip is not a zip archive"""
    with pytest.raises(Exception) as e_info:
        createFromZip("tests/blank.bin", "src/data", "./fail.xlam")
