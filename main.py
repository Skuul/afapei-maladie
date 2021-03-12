import os
from sys import path
import openpyxl
from algo import algo

# TODO: Make dirs

inDir = os.sep.join([os.getcwd(),'in'])
outDir = os.sep.join([os.getcwd(),'out'])

if not os.path.exists(outDir):
    os.mkdir(outDir)


files = os.listdir(inDir)
for file in files:
    if file[-5:] == '.xlsx':
        print(file)
        algo(os.sep.join([inDir,file]), outDir=outDir)
    else:
        print(f'Detected file "{file}" is not an xlsx file')
