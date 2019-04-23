import shutil
import sys
# check the extension on sys.argv[1] to determine which function to call
# if .xlam
#     createFromZip(sys.argv[1], sys.argv[2])
# else if .bin
#     createFromBin(sys.argv[1], sys.argv[2])

def createFromBin(input, output)
    # check that input is a binary file or the correct type
    shutil.move(input, "data/xl/vbaProject.bin")
    shutil.makearchive(output, 'zip', "data")

def createFromZip(input, output)
    # check that input is a zip file
    # check that the zip archive contains /xl/vbaProject.bin
    # binFile = {extract vbaProject.bin from input}
    shutil.move(binFile, "data/xl/vbaProject.bin")
    shutil.makearchive(output, 'zip', "data")

