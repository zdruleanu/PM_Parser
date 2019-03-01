import zipfile
from sys import argv
import os
from shutil import rmtree

fileName = argv[1]
compressedDataFolder = './compressedData'
dataFolder = './data'
targetFolder = os.path.splitext(fileName)[0]
targetPath = dataFolder + "/" + targetFolder


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
compressedFile = zipfile.ZipFile(compressedDataFolder + "/" + fileName)
print("Extracting files to temporary folder")
compressedFile.extractall(targetPathTmp)
print("done extracting tar.gz files")



