import shutil
import sys

if len(sys.argv) > 1:
    # check the extension on sys.argv[1] to determine which function to call
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    if input_file.endswith('.xlam'):
        createFromZip(input_file, output_file)
    elif input_file.endswith('.bin'):
        createFromBin(input_file, output_file)
    else:
        raise Exception(input_file . " is not a valid file format.")
        
def createFromBin(input_file, output_file)
    # check that input is a binary file or the correct type
    shutil.move(input_file, "data/xl/vbaProject.bin")
    shutil.makearchive(output_file, 'zip', "data")

def createFromZip(input_file, output_file)
    # check that input is a zip file
    # check that the zip archive contains /xl/vbaProject.bin
    # binFile = {extract vbaProject.bin from input}
    shutil.move(binFile, "data/xl/vbaProject.bin")
    shutil.makearchive(output_file, 'zip', "data")

