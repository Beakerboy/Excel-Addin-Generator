# test_excelAddinGenerator.py

import pytest
from excelAddinGenerator.main import *
from os.path import exists
from filehash import FileHash

def test_success_from_bin():
    """Test that xlam is successfully generated from a OLE file"""
    createFromBin("tests/vbaProject.bin", "src/data", "success_bin.xlam")
    # Assert that xlam file is created
    assert exists("success_bin.xlam")
    #assert that bin file within success_bin.xlam matches tests/vbaProject.bin
    extractBinFromZip("success_bin.xlam")
    assert hash_file("tests/vbaProject.bin") == hash_file("xl/vbaProject.bin")

    createFromZip("success_bin.xlam", "src/data", "success_xlam.xlam")
    assert exists("success_xlam.xlam")
    #ToDo: assert that bin file within success_xlam.xlam matches bin file within success_bin.xlam
    extractBinFromZip("success_xlam.xlam")
    assert hash_file("tests/vbaProject.bin") == hash_file("xl/vbaProject.bin")
    
def test_not_bin_exception():
    """ Test that an exception is thrown if the bin file is not an OLE file"""
    with pytest.raises(Exception) as e_info:
        createFromBin("tests/blank.bin", "src/data", "./fail.xlam")
        
def test_xlam_not_zip():
    """ Test that an exception is thrown if the zip is not a zip archive"""
    with pytest.raises(Exception) as e_info:
        createFromZip("tests/blank.bin", "src/data", "./fail.xlam")

def test_main():
    main(["./excelAddinGenerator", "./tests/vbaProject.bin", "success_bin.xlam"])
    main(["./excelAddinGenerator", "success_bin.xlam", "success_xlam.xlam"])

def test_main_incorrect_type():
    """ Test that an exception is thrown if the zip is not a zip archive"""
    with pytest.raises(Exception) as e_info:
        main(["./excelAddinGenerator", "./src/data/xl/styles.xml", "fail.xlam"])
