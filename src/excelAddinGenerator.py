import shutil
import sys
# Check that sys.argv[1] is a binary file or a zip archive with a binary file in the correct directory
shutil.move(sys.argv[1], "data/xl/vbaProject.bin")
shutil.makearchive(sys.argv[2], 'zip', "data")
