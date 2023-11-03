# test_excelAddinGenerator.py
import excel_addin_generator.main
import pytest
from os.path import exists
from filehash import FileHash

def test_success_from_bin():
    """Test that xlam is successfully generated from a OLE file"""
    excel_addin_generator.main.create_from_bin("tests/vbaProject.bin", "src/data", "success_bin.xlam")
    # Assert that xlam file is created
    assert exists("success_bin.xlam")
    #assert that bin file within success_bin.xlam matches tests/vbaProject.bin
    excel_addin_generator.main.extractBinFromZip("success_bin.xlam")
    md5hasher = FileHash('md5')
    assert md5hasher.hash_file("tests/vbaProject.bin") == md5hasher.hash_file("xl/vbaProject.bin")

    excel_addin_generator.main.createFromZip("success_bin.xlam", "src/data", "success_xlam.xlam")
    assert exists("success_xlam.xlam")
    #assert that bin file within success_xlam.xlam matches bin file within success_bin.xlam
    excel_addin_generator.main.extractBinFromZip("success_xlam.xlam")
    assert md5hasher.hash_file("tests/vbaProject.bin") == md5hasher.hash_file("xl/vbaProject.bin")
    
def test_not_bin_exception():
    """ Test that an exception is thrown if the bin file is not an OLE file"""
    with pytest.raises(Exception) as e_info:
        excel_addin_generator.main.createFromBin("tests/blank.bin", "src/data", "./fail.xlam")
        
def test_xlam_not_zip():
    """ Test that an exception is thrown if the zip is not a zip archive"""
    with pytest.raises(Exception) as e_info:
        excel_addin_generator.main.createFromZip("tests/blank.bin", "src/data", "./fail.xlam")

def test_main():
    excel_addin_generator.main.main(["./excelAddinGenerator", "./tests/vbaProject.bin", "success_bin.xlam"])
    excel_addin_generator.main.main(["./excelAddinGenerator", "success_bin.xlam", "success_xlam.xlam"])

def test_main_incorrect_type():
    """ Test that an exception is thrown if the zip is not a zip archive"""
    with pytest.raises(Exception) as e_info:
        excel_addin_generator.main.main(["./excelAddinGenerator", "./src/data/xl/styles.xml", "fail.xlam"])
