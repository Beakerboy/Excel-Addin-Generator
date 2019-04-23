import shutil
import sys
shutil.move(sys.argv[1], "data/xl/vbaProject.bin")
shutil.makearchive(sys.argv[2], 'zip', "data")
