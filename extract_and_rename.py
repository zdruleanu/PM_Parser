import zipfile
import tarfile
from sys import argv
import os
from shutil import rmtree


def folder_exists_menu(target_path):
    # Used to create a menu and return the target path which is used to extract and process the data
    # if the user chooses just to process, it retunrs the already existing target path.
    # Therefore the data processed will be the already existing data in the folder

    print('Folder ', target_path, ' already exists. Replace folder contents or create a new folder: ')
    while True:
        print('1. Replace folder contants')
        print('2. Create a new folder')
        print('3. Exit')
        selection = input('Your choice: ')
        if selection == '1':
            print("Removing the target folder and files")
            rmtree(target_path)
            targetPath = target_path
            break
        elif selection == '2':
            targetFolder = input('Please input a new destination folder name: ')
            targetPath = os.path.split(target_path)[0] + "/" + targetFolder
            if os.path.isdir(targetPath):
                print('Folder ', targetPath, ' also already exists. Please make a valid choice for ', target_path)
            else:
                break
        elif selection == '3':
            print("exiting program")
            exit()
        else:
            print("Please make a valid choice")
    return targetPath


compressedDataFolder = './compressedData'

try:
    zipFileName = argv[1]
except IndexError as e:
    print("Zip file name not provided as argument")
    zipFileName = ""
while True:
    if os.path.isfile(compressedDataFolder + "/" + zipFileName):
        print("Zip file containing data: " + zipFileName)
        break
    else:
        print("zip file doesn't exist in " + compressedDataFolder)
        zipFileName = input('Please enter the zip filename: ')
dataFolder = './data'
targetFolder = os.path.splitext(zipFileName)[0]
targetPath = dataFolder + "/" + targetFolder

if not (os.path.isdir(targetPath)):
    os.mkdir(targetPath)
else:
    # We ask the user what to do about the existing folder
    targetPath = folder_exists_menu(targetPath)
    print("Creating the target folder: ", targetPath)
    os.mkdir(targetPath)

# targetPathTmp is used to extract from the zip file, the tar.gz files.
# These files will then be extracted to the targetPath
targetPathTmp = targetPath + "/tmp"

if os.path.isdir(targetPathTmp):
    print("temporary folder containing tar.gz files exists. Removing the folder and contents")
    rmtree(targetPathTmp)

print("Creating the temporary folder for tar.gz files")
os.mkdir(targetPathTmp)
compressedFile = zipfile.ZipFile(compressedDataFolder + "/" + zipFileName)

# orderFile contains the order of the extracted files. It is actually the prefixes which will be used for the filenames
orderFilePath = compressedDataFolder + "/" + zipFileName + "_order.txt"
with open(orderFilePath,"r") as orderFile:
    orderList = orderFile.readlines()
print("Extracting files to temporary folder")
compressedFile.extractall(targetPathTmp)
print("done extracting from zip file")
targzfFiles = [targetPathTmp + "/" + filename for filename in os.listdir(targetPathTmp)
               if os.path.isfile(targetPathTmp + "/" + filename)]

print("Extracting and renaming PM files")
# the elements in orderList are assumed to match the order of the files extracted from the zip archive
for (fileName, order) in zip(targzfFiles, orderList):
    targzFile = tarfile.open(fileName)
    [pmFileNameTarInfoObj] = targzFile.getmembers()
    pmFileName = pmFileNameTarInfoObj.name
    targzFile.extractall(path=targetPath)
    # we only use order[:-1] in order to demove the \n
    targetPmFileName = order[:-1] + "_" + pmFileName.split("_")[-1]
    os.rename(targetPath + "/" + pmFileName, targetPath + "/" + targetPmFileName)
    print("Extracted " + pmFileName + " as " + targetPmFileName)