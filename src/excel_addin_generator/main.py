import os
import shutil
import sys
import zipfile


 create_from_bin(input_file: str, wrapper_dir: str, output_file_name: str) -> None:

    """Create a zip file containing the provided bin"""
    # file must start with 'd0 cf 11 e0 a1 b1 1a e1'
    fileSig = open(input_file, "rb").read(8).hex()
    if fileSig != 'd0cf11e0a1b11ae1':
        raise Exception('File signature {} is not as expected.', format(fileSig))
    shutil.copy(input_file, wrapper_dir + "/xl/vbaProject.bin")
    shutil.make_archive(output_file_name, 'zip', wrapper_dir)
    shutil.move(output_file_name + ".zip", output_file_name)


def createFromZip(input_file: str, wrapper_dir: str, output_file_name: str) -> None:

    """Create a zip file containing the bin file within the provided zip file"""
    extractBinFromZip(input_file)
    create_from_bin('xl/vbaProject.bin', wrapper_dir, output_file_name)


def extractBinFromZip(input_file: str) -> None:

    # check that input is a zip file
    if zipfile.is_zipfile(input_file):
        # check that the zip archive contains /xl/vbaProject.bin
        with zipfile.ZipFile(input_file, 'r') as zip:
            zip.extract('xl/vbaProject.bin')
    else:
        raise Exception(input_file, " is not a valid file format.")
