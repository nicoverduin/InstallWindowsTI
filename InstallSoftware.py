#
# program   : ModifyPath
# purpose   : remove any bmptk references from the Windows PATH variable
# author    : Nico Verduin
# date      : 16-5-2020
#
import winreg
import os.path
import subprocess
import urllib.request
import pathlib
import logging
import sys
import struct
import shutil

# determine environment
print(struct.calcsize("P") * 8)


# Create our logger
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s.%(msecs)03d %(levelname)s %(module)s : %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S',
                    filename='Install.log',
                    filemode='w')
logger = logging.getLogger()
# copy logging output to stdout
logger.addHandler(logging.StreamHandler(sys.stdout))

# get current working folder
CWD = pathlib.Path().absolute().__str__() + "\\"

logger.info("Start installation")

# Prerequisites
logger.info(("Verifying 7zip and git are installed"))

# no need to find python as this program would not work

# find 7z. This could be de 32 bit or 64 bit version
ZipProgram = ""
zipLocations = ["C:\\Program Files\\7-Zip\\7z.exe", "C:\\Program Files (x86)\\7-Zip\\7z.exe"]
for file in zipLocations:
    if os.path.exists(file):
        ZipProgram = file

if ZipProgram == "":
    # We cannot continue
    logger.info("Cannot find installation of 7z. Cancelling installation")
    logger.info("Please install 7z from hhttps://www.7-zip.org/download.html")
    exit(1)
else:
    logger.info("found 7z program : " + ZipProgram)

# existence of git program
GitProgram = ""
gitLocations = ["C:\\Program Files\\Git\\cmd\\git.exe", "C:\\Program Files (x86)\\Git\\cmd\\git.exe"]
for file in gitLocations:
    if os.path.exists(file):
        GitProgram = file

if GitProgram == "":
    # We cannot continue
    logger.info("Cannot find installation of GIT. Cancelling installation")
    logger.info("Please install GIT desktop from https:/desktop.github.com/")
    exit(1)
else:
    logger.info("found GIT program : " + GitProgram)


logger.info("Build GIT list")
GitLocations    = [ "https://github.com/wovo/bmptk.git",
                    "https://github.com/wovo/hwlib.git",
                    "https://github.com/wovo/rtos.git",
                    "https://github.com/wovo/v1oopc-examples.git",
                    "https://github.com/wovo/v2cpse1-examples.git",
                    "https://github.com/wovo/v2cpse2-examples.git",
                    "https://github.com/wovo/v2thde-examples.git",
                    "https://github.com/catchorg/Catch2.git"]

logger.info("Build Compiler list")
Compilers    =  [["GCC-ARM", "https://developer.arm.com/-/media/Files/downloads/gnu-rm/9-2019q4/gcc-arm-none-eabi-9-2019-q4-major-win32.zip"],
                 ["GCC-WIN", "http://ftp.vim.org/languages/qt/development_releases/prebuilt/mingw_32/i686-7.3.0-release-posix-dwarf-rt_v5-rev0.7z"],
                 ["GCC-AVR", "https://blog.zakkemble.net/download/avr-gcc-9.2.0-x86-mingw.zip"],
                 ["SFML", "https://www.sfml-dev.org/files/SFML-2.5.1-windows-gcc-7.3.0-mingw-32-bit.zip"]]


# download an install our GIT repositories
logger.info("Downloading GIT repositories")
for repo in GitLocations:
    repoName = repo[repo.rfind("/")+1:repo.rfind(".")]
    if os.path.exists(repoName):
        logger.info("Skipping git clone. Repository " + repoName + " exists and/or is not empty.")
    else:
        process = subprocess.run(["git", "clone", repo, repoName], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        data =  str(process.stdout)
        logger.info(data[2:-3])

logger.info("Downloading Compilers")
for compiler in Compilers:
    # get our local file name
    compilerFile = compiler[1][compiler[1].rfind("/")+1:]

    # add column in compiler with our BIN path name for the custom make later on
    makefileName = compilerFile[:compilerFile.rfind(".")]
    compiler.append(makefileName)

     # check if the zip file is already there
    if os.path.exists(compilerFile):
        logger.info("skipping download : " + compilerFile + " as it already exists")
        decompressable = True
    else:
        # not here yet, so download it
        logger.info("Downloading from : " + compiler[1] + " into " + compilerFile)
        # get the filesize to be downloaded
        req = urllib.request.Request(compiler[1], method='HEAD')
        f = urllib.request.urlopen(req)
        fileSize = int(f.headers['Content-Length'])

        # start download
        urllib.request.urlretrieve(compiler[1], compilerFile)

        # check if filesizes are equal
        if os.stat(compilerFile).st_size == fileSize:
            logger.info("Download " + compilerFile + " was successful")
            decompressable = True
        else:
            logger.info("Download " + compilerFile + " failed. Downloaded " + os.stat(compilerFile).st_size + " should be : " + fileSize)
            decompressable = False

    # if it seems decompressable, may as well do it
    if decompressable:
        # extract in root of compilername
        foldername = compilerFile[0:compilerFile.rfind(".")]
        logger.info("Decompressing " + compilerFile + "...")
        #
        # TODO: see if finding install directory is better choice
        #
        if foldername.find("SFML-2.5.1") >= 0:
            # delete any old SFML-2.5.1-32 folder
            logger.info("Delete any existing SFML-2.5.1-32 folder")
            shutil.rmtree('SFML-2.5.1', ignore_errors=True)
            process = subprocess.run(
                [ZipProgram, "x", compilerFile, "SFML-2.5.1", "-o.", "-y"],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

            # we get more output lines but need some cleaning
            data = str(process.stdout)[2:-3].replace('\\r', '').replace('\\n', '\n').splitlines()
            for line in data:
                logger.info(line)

            # rename the folder to SFML-2.5.1-32
            # logger.info("Rename SFML-2.5.1 to SFML-2.5.1-32")
            # os.rename("SFML-2.5.1", "SFML-2.5.1-32")
        else:
            process = subprocess.run(
                [ZipProgram, "x", compilerFile, "-o"+foldername, "-y"],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            data = str(process.stdout)[2:-3].replace('\\r', '').replace('\\n', '\n').splitlines()
            for line in data:
                logger.info(line)

# Special operation for the windows GCC-WIN compiler
for compiler in Compilers:
    if compiler[0] == "GCC-WIN":
        # we have to check if it is the 32 or 64 bit compiler
        if os.path.exists(compiler[2] + "\\mingw32"):
            compiler[2] = compiler[2] +  "\\mingw32"
        else:
            compiler[2] = compiler[2] +  "\\mingw64"

# create Makefile.custom in bmptk
makefile = open("bmptk\\Makefile.local", "r")
custom   = open("bmptk\\Makefile.custom", "w")

logger.info("Building Makefile.custom")
windowsPart = False

makeLines = makefile.read().splitlines()
for line in makeLines:
    outputLine = line
    line = line.strip()
    # check if this is the Windows part
    if line.find("ifeq ($(OS),Windows_NT") >= 0:
        windowsPart = True
# only process the Windows part
    if windowsPart:
        if line.find("else") >= 0:
            windowsPart = False
        else:
            # do nothing with comment lines
            if len(line) != 0:
                if line[0] != "#":
                    # scan if it is a compiler definition
                    for compiler in Compilers:
                        searchArgument = compiler[0]
                        if (len(line) >= len(searchArgument)):
                            if line[0:len(searchArgument)] == compiler[0]:
                                outputLine = "   " + compiler[0] + "          ?= ..\\..\\" + compiler[2]

    custom.write(outputLine + "\n")

makefile.close()
custom.close()
logger.info("Building Makefile.custom completed")

# *** Adjusting the PATH ***

logger.info("Update PATH in Windows Registry")
pathSeparator   = ";"
regEditFile     = "PathUpdate.reg"
pathNeedsUpdate = False


logger.info("Read current PATH in Windows Registry")
# Read the registry for our system path variable
root_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment")
[Pathname,regtype]=(winreg.QueryValueEx(root_key,"Path"))
winreg.CloseKey(root_key)

# change it into a list
Paths = Pathname.split(";")

logger.info("Remove any references to BMPTK")
# remove any path that refers to \BMPTK\
for str in Paths:
    y = str.upper()
    if y.find("\\BMPTK\\") >= 0:
        pathNeedsUpdate = True
        logger.info("Deleted " + str)
        Paths.remove(str)

logger.info("Add new BMPTK path : " + CWD + "bmptk\\tools")
# rebuild our path string
Pathname = CWD + 'bmptk\\tools' + pathSeparator
for str in Paths:
    Pathname += str + pathSeparator

# Change a single backslash to double (special character)
convertedPathname = Pathname.replace('\\', '\\\\')

logger.info("Create .REG file")
# build a new registration file
outputFile = open(regEditFile, "w")
outputFile.write("Windows Registry Editor Version 5.00\n\n")
outputFile.write("[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment]\n")

# delete old Path value
outputFile.write('\"Path\"=-\n')

# add our new Path parameter
outputFile.write('\"Path\"=\"' + convertedPathname + '\"\n\n')
outputFile.close()

# build our REGEDIT command
regFile = CWD + regEditFile

# execute sub processes as administrator

logger.info("Update Windows Registry")
# execute regedit
subprocess.run(["runas", "/noprofile", "/user:Administrator", "|", "echo", "Y", "|", "regedit","/s", regFile], shell=True)

logger.info("Delete .REG file")
# remove the generated registration file
subprocess.run(["runas","/noprofile", "/user:Administrator","|", "echo", "Y", "|","del", regFile], shell=True)


# prepare example foldes for codelite

logger.info("processing codelite update v1oopc-examples")
os.chdir("v1oopc-examples")
process = subprocess.run(["Python.exe", "./../bmptk/tools/bmptk-mef.py", "-os",  "windows"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

data = process.stdout.splitlines()
for line in data:
    logger.info(line)
    
logger.info("processing codelite update v2cpse1-examples")
os.chdir("../v2cpse1-examples")
process = subprocess.run(["Python.exe", "./../bmptk/tools/bmptk-mef.py", "-os",  "windows"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

data = process.stdout.splitlines()
for line in data:
    logger.info(line)

logger.info("processing codelite update v2cpse2-examples")
os.chdir("../v2cpse2-examples")
process = subprocess.run(["Python.exe", "./../bmptk/tools/bmptk-mef.py", "-os",  "windows"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

data = process.stdout.splitlines()
for line in data:
    logger.info(line)

logger.info("processing codelite update v2thde-examples")
os.chdir("../v2thde-examples")
process = subprocess.run(["Python.exe", "./../bmptk/tools/bmptk-mef.py", "-os",  "windows"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

data = process.stdout.splitlines()
for line in data:
    logger.info(line)

# done 

logger.info("Installation completed")
logger.info("\n*****   Please Reboot the computer to make the PATH change effective!! *****")
