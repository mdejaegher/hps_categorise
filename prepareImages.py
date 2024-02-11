#!/usr/bin/python
from CImageInfo import CImageInfo
from CImageInfo import CPendingImageInfo
from CSpreadSheet import CSpreadSheet

import argparse
import datetime
import os
import re
import shutil
import subprocess
import sys

# Force print to always flush
import functools
print = functools.partial(print, flush=True)


class CHPS:

    # Column names with numbers in RHS database
    RHS_OLDSPECIESCODE = 1
    RHS_CALCTOPRANKEDENTITYNAME = 2
    RHS_CACLFULLNAME = 3
    RHS_FAMILYNAME = 4
    RHS_GENUSNAME = 5
    RHS_SPECIESNAME = 6
    RHS_CULTIVAR = 13

    def __init__(self, args):
        self.args = args

        self.scriptDir = 'H:\\hps_categorise\\'
        self.baseDir = 'H:\\HPS_Images\\'
        self.plantsDir = self.baseDir + 'Plants\\'
        self.pendingPlantsDir = self.baseDir + 'Pending\\Plants\\'
        self.gardensDir = self.baseDir + 'Gardens\\'
        self.pendingGardensDir = self.baseDir + 'Pending\\Gardens\\'
        self.thumbsDir = self.baseDir + 'Thumbnails\\'
        self.uploadDir = self.baseDir + 'Upload_'+datetime.datetime.now().strftime("%d%m%y")+'\\'

        # Current data
        self.hpsPlantsImageInfo = []
        self.hpsGardensImageInfo = []
        # Pending data
        self.pendingPlantImages = True
        self.pendingPlantsImageInfo = []
        self.pendingGardenImages = True
        self.pendingGardensImageInfo = []
        # Upload directories
        self.uploadPlantsDir = self.uploadDir+'Plants\\'
        self.uploadGardensDir = self.uploadDir+'Gardens\\'
        self.uploadThumbsDir = self.uploadDir+'thumbs\\'
        self.uploadUnknownProvenancePlantsDir = self.uploadDir+'unknownProvenance\\'

    def validateDirectories(self):
        print("* Validate directories")

        # Check the base directory
        print(f"  - Base directory, '{self.baseDir}': ", end="")
        if not os.path.isdir(self.baseDir):
            print("doesn't exist!")
            return 1
        print("OK")

        # Check the directory containing all the current plants images. Should
        # have a directory for each letter of the alphabet
        print(f"  - HPS plants images directory, {self.plantsDir}: ", end="")
        if not os.path.isdir(self.plantsDir):
            print("doesn't exist!")
            return 1
        hpsPlantsLetterDirs = []
        if os.listdir(self.plantsDir):
            for startletter in os.listdir(self.plantsDir):
                hpsPlantsLetterDirs.append(self.plantsDir+startletter)
        if len(hpsPlantsLetterDirs) == 0:
            print("can't find any plant directories!")
            return 1
        print("OK")

        # Check the directory containing the pending plants images.
        print(f"  - Pending plants images directory, '{self.pendingPlantsDir}': ")
        if not os.path.isdir(self.pendingPlantsDir):
            print("    - doesn't exist!")
            self.pendingPlantImages = False
        else:
            print("    - exists")
            if not os.listdir(self.pendingPlantsDir):
                print("    - is empty")
                self.pendingPlantImages = False
            else:
                print("    - contains images")

        # Check the directory containing all the current garden images.
        print(f"  - HPS gardens images directory, {self.gardensDir}: ", end="")
        if not os.path.isdir(self.gardensDir):
            print("doesn't exist!")
            return 1
        print("OK")

        # Check the directory containing the pending gardens images.
        print(f"  - Pending gardens images directory, '{self.pendingGardensDir}': ")
        if not os.path.isdir(self.pendingGardensDir):
            print("    - doesn't exist!")
            self.pendingGardenImages = False
        else:
            print("    - exists")
            if not os.listdir(self.pendingGardensDir):
                print("    - is empty")
                self.pendingGardenImages = False
            else:
                print("    - contains images")

        # Check the directory containing all the current thumbnails.
        print(f"  - HPS thumbnails directory, {self.thumbsDir}: ", end="")
        if not os.path.isdir(self.thumbsDir):
            print("doesn't exist!")
            return 1
        print("OK")

        # No point continuing if neither pending plants or pending garden images
        # are available
        if not self.pendingPlantImages and not self.pendingGardenImages:
            print("! Can't find any pending plants or gardens. Exit")
            return 1

        # If the upload directory doesn't exist, create it
        print(f"  - Upload directory, '{self.uploadDir}': ", end="")
        if not os.path.isdir(self.uploadDir) and not self.args.dryrun:
            try:
                os.mkdir(self.uploadDir)
            except OSError:
                print("couldn't create directory!")
                return 1
        print("OK")

        return 0

    def createImagelibDB(self):
        # imagelib.csv can be found in docsftp@hardy-plant.org.uk:/plants
        fileName = self.scriptDir+"imagelib.csv"
        print(f"  - {fileName}: importing ", end="\r")
        if not os.path.isfile(fileName):
            print("doesn't exist!")
            return 1
        self.imagelibDB = CSpreadSheet(fileName)
        print(f"  - {fileName}: OK         ")

        return 0

    def createGeneraDB(self):
        # genera.csv can be found in docsftp@hardy-plant.org.uk:/plants
        fileName = self.scriptDir+"genera.csv"
        print(f"  - {fileName}: importing ", end="\r")
        if not os.path.isfile(fileName):
            print("doesn't exist!")
            return 1
        self.generaDB = CSpreadSheet(fileName)
        print(f"  - {fileName}: OK         ")

        return 0

    def createHpsPlantsDB(self):
        # Get the latest version of the plants database. This needs to be done
        # better by checking if files are same or not using requests
        fileName = self.scriptDir+"HPS Images - Plants.xlsx"
        print(f"  - {fileName}: importing  ", end="\r")
        self.hpsPlantsDB = CSpreadSheet(fileName)
        print(f"  - {fileName}: OK         ")

        return 0

    def createHpsGardensDB(self):
        # Get the latest version of the gardens database. This needs to be done
        # better by checking if files are same or not using requests
        fileName = self.scriptDir+"HPS Images - Gardens.xlsx"
        print(f"  - {fileName}: importing  ", end="\r")
        self.hpsGardensDB = CSpreadSheet(fileName)
        print(f"  - {fileName}: OK         ")

        return 0

    def checkConsistency(self):
        # Find where the P files (plants) change into X files (gardens)
        imagelibAccession = ""
        for index in range(2, self.imagelibDB.workbook['active'].max_row):
            imageID = self.imagelibDB.getValue('active', index, 2)  # Image ID
            if imageID.startswith("X"):
                break
            imagelibAccession = imageID

        if self.pendingPlantImages:
            maxRow = self.hpsPlantsDB.workbook['Plants'].max_row
            hpsPlantsAccession = self.hpsPlantsDB.getValue('Plants', maxRow, 2)  # Number

            if imagelibAccession != hpsPlantsAccession:
                print(f"* Libraries not consistent! {imagelibAccession} vs {hpsPlantsAccession}")
                return 1

        if self.pendingGardenImages:
            imagelibAccession = self.imagelibDB.getValue('active', self.imagelibDB.workbook['active'].max_row, 2)  # Image ID
            maxRow = self.hpsGardensDB.workbook['Gardens'].max_row
            hpsGardensAccession = self.hpsGardensDB.getValue('Gardens', maxRow, 2)  # Number

            if imagelibAccession != hpsGardensAccession:
                print(f"* Libraries not consistent! {imagelibAccession} vs {hpsGardensAccession}")
                return 1

        return 0

    def createRhsReferenceDB(self):
        # Get the latest version of the RHS database. This needs to be done better
        # by checking if files are same or not using requests
        fileName = self.scriptDir+"RHS_0923_Reduced_Unlocked.xlsx"
        print(f"  - {fileName}: importing  ", end="\r")
        self.rhsReferenceDB = CSpreadSheet(fileName)
        print(f"  - {fileName}: OK         ")

        return 0

    def validateDatabases(self):
        print("* Validate and import databases")
        # Import imagelib.csv which is a number ordered list of all the plant
        # and garden images on the website
        # For plants, genus+species is in italic, cultivar is normal (should be
        # same as in NAME_HTML in rhs dataset
        if self.createImagelibDB():
            return 1
        expectedHeaders = ["Caption",
                           "Image ID"]
        if self.imagelibDB.validate('active', expectedHeaders):
            return 1

        # Import genera.csv which is an alphabetically sorted list of
        # genus/family plant names in database.
        # Used by website to create menus and titles. Comments are shown on
        # genus pages under the title
        if self.pendingPlantImages:
            if self.createGeneraDB():
                return 1
            expectedHeaders = ["genus",
                               "family",
                               "notes"]
            if self.generaDB.validate('active', expectedHeaders):
                return 1

        # Import 'HPS Images - Plants.xlsx' which is the central HPS
        # plants database
        if self.pendingPlantImages:
            if self.createHpsPlantsDB():
                return 1
            expectedHeaders = ["Plant name",    "Number",            "RHS no",
                               "RHS status",    "qualifier",         "descriptor",
                               "image caption", "Donor",             "Date added",
                               "Slide No.",     "Extra information", "Date withdrawn"]
            if self.hpsPlantsDB.validate('Plants', expectedHeaders):
                return 1

        # Import 'HPS Images - Gardens.xlsx' which is the central
        # HPS gardens database
        if self.pendingGardenImages:
            if self.createHpsGardensDB():
                return 1
            expectedHeaders = ["Topic",         "Number",    "Donor",
                               "Date added",    "Slide No.", "Extra Information",
                               "Date withdrawn"]
            if self.hpsGardensDB.validate('Gardens', expectedHeaders):
                return 1

        # Check consistency between HPS and imagelib datasets
        if self.checkConsistency():
            return 1

        # Import RHS_Dataset.xlsx which is the RHS dataset
        if self.pendingPlantImages:
            if self.createRhsReferenceDB():
                return 1
            expectedHeaders = ["OldSpeciesCode", "CalcTopRankedEntityName", "CalcFullName",
                               "FamilyName", "GenusName", "SpeciesName", "Subspecies", "Variety", "Subvariety",
                               "Forma", "TradeSeries", "TradeDesignation", "Cultivar", "Descriptor"]
            if self.rhsReferenceDB.validate('Table1', expectedHeaders):
                return 1

        return 0

    def validateTools(self):
        print("Validate tools")
        print("--------------")

        # Check if 'magick' in path
        if shutil.which('magick') is None:
            print("! Can't find 'magick' in path. This is needed to resize and add watermarks to images.")
            print("! To download, do following steps:")
            print("!   * go to 'https://imagemagick.org/script/download.php'")
            print("!   * download the executable (e.g. 'ImageMagick-[version]-Q16-HDRI-x64-dll.exe')")
            print("!   * run the executable and it should be installed for you")
            return 1
        else:
            print("* Found 'magick'")

        # Check if 'exiftool' in path
        if shutil.which('exiftool') is None:
            print("! Can't find 'exiftool' in path. This is needed to removed some exif data in images for security.")
            print("! To download, do following steps:")
            print("!   * go to 'https://exiftool.org/'")
            print("!   * download the zip file")
            print("!   * unzip to file into an appropriate directory")
            print("!   * rename to exiftool.exe")
            print("!   * add to PATH (see 'Environment Variables' in 'System Properties'")
            return 1
        else:
            print("* Found 'exiftool'")

        print()
        return 0

    def validateInput(self):
        print("Validate input")
        print("--------------")

        # Validate the directories
        if self.validateDirectories():
            return 1

        # Validate the xlsx files
        if self.validateDatabases():
            return 1

        print()
        return 0

    def importCurrentImages(self):
        if self.pendingPlantImages:
            for plantsLetterDir in os.listdir(self.plantsDir):
                print(f"  - directory '{plantsLetterDir}'", end="\r")
                for filename in os.listdir(self.plantsDir+plantsLetterDir):
                    fullpath = self.plantsDir + plantsLetterDir + '\\' + filename
                    self.hpsPlantsImageInfo.append(CImageInfo(fullpath, False))
            print(f"  - Imported current plant images{' ': <108}")

        if self.pendingGardenImages:
            for filename in os.listdir(self.gardensDir):
                fullpath = self.gardensDir + '\\' + filename
                self.hpsGardensImageInfo.append(CImageInfo(fullpath, False))
            print(f"  - Imported current garden images{' ': <108}")

    def importPendingImages(self):
        if self.pendingPlantImages:
            for filename in os.listdir(self.pendingPlantsDir):
                fullpath = self.pendingPlantsDir + filename
                print(f"  - {filename: <108}", end="\r")
                self.pendingPlantsImageInfo.append(CPendingImageInfo(fullpath))
            print(f"  - Imported pending plant images{' ': <108}")
        if self.pendingGardenImages:
            for filename in os.listdir(self.pendingGardensDir):
                fullpath = self.pendingGardensDir + filename
                print(f"  - {filename: <108}", end="\r")
                self.pendingGardensImageInfo.append(CPendingImageInfo(fullpath))
            print(f"  - Imported pending garden images{' ': <108}")

    def getImageInfo(self):
        print("* Import existing images")
        self.importCurrentImages()
        if self.pendingPlantImages and len(self.hpsPlantsImageInfo) == 0:
            print("! Couldn't find information for current plant images\n")
            return 1
        if self.pendingGardenImages and len(self.hpsGardensImageInfo) == 0:
            print("! Couldn't find information for current garden images\n")
            return 1

        print("* Import pending images")
        self.importPendingImages()
        if self.pendingPlantImages and len(self.pendingPlantsImageInfo) == 0:
            print("! Couldn't find information for pending plant images\n")
            return 1
        if self.pendingGardenImages and len(self.pendingGardensImageInfo) == 0:
            print("! Couldn't find information for pending garden images\n")
            return 1

        return 0

    def importImages(self):
        print("Import images")
        print("-------------")

        # Get file information for pending and existing files
        if self.getImageInfo():
            return 1

        print()
        return 0

    def constainsName(self, shortName, longName):
        name1 = shortName.lower()
        name1 = self.convertSpecialChar(name1)
        name1 = name1.replace(' ', '')
        name1 = name1.replace('[', '')
        name1 = name1.replace(']', '')
        name1 = name1.replace("'", "")

        name2 = longName.lower()
        name2 = self.convertSpecialChar(name2)
        name2 = name2.replace(' ', '')
        name2 = name2.replace('[', '')
        name2 = name2.replace(']', '')
        name2 = name2.replace("'", "")

        if name1 in name2:
            return True
        return False

    def convertSpecialChar(self, value):
        if not value:
            return value

        value = value.replace(u'\N{LATIN CAPITAL LETTER A WITH DIAERESIS}',  u'A')
        value = value.replace(u'\N{LATIN CAPITAL LETTER E WITH GRAVE}',      u'E')
        value = value.replace(u'\N{LATIN CAPITAL LETTER E WITH ACUTE}',      u'E')
        value = value.replace(u'\N{LATIN CAPITAL LETTER N WITH TILDE}',      u'N')
        value = value.replace(u'\N{LATIN CAPITAL LETTER O WITH DIAERESIS}',  u'O')
        value = value.replace(u'\N{LATIN CAPITAL LETTER O WITH CIRCUMFLEX}', u'O')
        value = value.replace(u'\N{LATIN CAPITAL LETTER U WITH DIAERESIS}',  u'U')
        value = value.replace(u'\N{LATIN CAPITAL LETTER U WITH CIRCUMFLEX}', u'U')

        value = value.replace(u'\N{LATIN SMALL LETTER A WITH DIAERESIS}',  u'a')  # E4
        value = value.replace(u'\N{LATIN SMALL LETTER E WITH GRAVE}',      u'e')  # E8
        value = value.replace(u'\N{LATIN SMALL LETTER E WITH ACUTE}',      u'e')  # E9
        value = value.replace(u'\N{LATIN SMALL LETTER N WITH TILDE}',      u'n')
        value = value.replace(u'\N{LATIN SMALL LETTER O WITH DIAERESIS}',  u'o')
        value = value.replace(u'\N{LATIN SMALL LETTER O WITH CIRCUMFLEX}', u'o')
        value = value.replace(u'\N{LATIN SMALL LETTER U WITH DIAERESIS}',  u'u')
        value = value.replace(u'\N{LATIN SMALL LETTER U WITH CIRCUMFLEX}', u'u')

        value = value.replace(u'\N{MULTIPLICATION SIGN}',                  u'x')
        value = value.replace(u'/',                                        u'_')

        return value

    def createHtmlTag(self, html):
        html = html.replace(u'[', '<span class="trade-name">')
        html = html.replace(u']', '</span>')
        html = "<i>" + html + "</i>"
        return html

    def createHtmlName(self, index):
        # This method wasn't needed in the previous database as the html was included in the
        # NAME_HTML column. This column doesn't exist any longer in the new database so we're
        # having to make it up.
        # I tried at first to reassemble it based on the indiviual component columns (e.g.
        # subspecies/variety/subvariety/... ) but that didn't work as there were too many
        # exceptions (e.g. any names with an ' x ' wasn't visible in any columns). I therefore
        # had to change it to do inline replacement.
        calcfullname = self.rhsReferenceDB.getValue('Table1', index, self.RHS_CACLFULLNAME)
        genusname = self.rhsReferenceDB.getValue('Table1', index, self.RHS_GENUSNAME)
        speciesname = self.rhsReferenceDB.getValue('Table1', index, self.RHS_SPECIESNAME)

        html = calcfullname

        # Search for the 'genusname speciesname' string. There are a few cases where there are
        # extra words in between which are not supposed to in italic.
        basicsearch = genusname + r" (\S )?(aff\. )?(.* gx )?"
        if speciesname:
            basicsearch += speciesname
        genspec = re.search(basicsearch, html)
        if not genspec:
            print(f"Can break down name '{calcfullname}' due to not conforming to '{basicsearch}'")
            return calcfullname

        # Genus name is always in italic
        src = genusname + " "
        dst = "<i>" + genusname + "</i> "
        if genspec.group(1):
            src += genspec.group(1)
            dst += genspec.group(1)
        if genspec.group(2):
            src += genspec.group(2)
            dst += genspec.group(2)
        if genspec.group(3):
            src += genspec.group(3)
            dst += genspec.group(3)
        # Speciesname (if present) is always in italic
        if speciesname:
            src += speciesname
            dst += "<i>" + speciesname + "</i>"
        # Create the correct html for 'genusname speciesname'
        html = html.replace(src, dst)

        # Subspecies is in italic
        subspsearch = r" subsp. (\S*)"
        subsp = re.search(subspsearch, html)
        if subsp:
            html = html.replace(" subsp. "+subsp.group(1), " subsp. <i>"+subsp.group(1)+"</i>")

        # Variety is in italic
        varsearch = r" var. (\S*)"
        var = re.search(varsearch, html)
        if var:
            html = html.replace(" var. "+var.group(1), " var. <i>"+var.group(1)+"</i>")

        # Subvariety is in italic
        subvarsearch = r" subvar. (\S*)"
        subvar = re.search(subvarsearch, html)
        if subvar:
            html = html.replace(" subvar. "+subvar.group(1), " subvar. <i>"+subvar.group(1)+"</i>")

        # Forma is in italic
        formasearch = r" f. (\S*)"
        forma = re.search(formasearch, html)
        if forma:
            html = html.replace(" f. "+forma.group(1), " f. <i>"+forma.group(1)+"</i>")

        # Replace any 'x' with html readable string
        html = html.replace(u'\N{MULTIPLICATION SIGN}', "&times;")

        return html

    def updateImageInfo(self):
        print("Update image info")
        print("-----------------")
        self.updatePlantImageInfo()
        self.updateGardenImageInfo()

    def updateGardenImageInfo(self):
        if not self.pendingGardenImages:
            return 0
        print("* Update garden images")
        for imageNum, imageInfo in enumerate(self.pendingGardensImageInfo):
            if imageInfo.valid is False:
                continue

            print(f"  - {imageNum+1}/{len(self.pendingGardensImageInfo)}: '{imageInfo.filename}'")
            imageData = re.search(r'(\D+)\s+\d+\s+(\D+)\s*(\d*)', imageInfo.filename.strip())
            if imageData:
                # Extract garden name
                imageInfo.gardenName = imageData.group(1)
                # Extract donor name
                imageInfo.donor = imageData.group(2).rstrip()
                # Extract date added
                if len(imageData.groups()) > 2 and imageData.group(3) != '0':
                    dateAdded = f"01/01/{imageData.group(3)}"
                # Extract information
                if len(imageData.groups()) > 3:
                    imageInfo.metaData = imageData.group(4)
            else:
                print("      ! File name doesn't conform to '<garden> <number> <donor> <year>' format")
                imageInfo.valid = False
                continue

            print(f"    - Got garden name extracted as '{imageInfo.gardenName}'")
            print(f"    - Got donor as '{imageInfo.donor}'")

            # If no date was extracted from the file name then take current date
            if not dateAdded:
                now = datetime.datetime.now()
                dateAdded = now.strftime("%d/%m/%Y")
            imageInfo.dateAdded = dateAdded
            print(f"    - Got date added as {imageInfo.dateAdded}")

        print()
        return 0

    def updatePlantImageInfo(self):
        if not self.pendingPlantImages:
            return 0

        print("Get RHS number (comma separated, empty to ignore)")
        print("-------------------------------------------------")
        newPlants = 0
        for imageNum, imageInfo in enumerate(self.pendingPlantsImageInfo):
            if imageInfo.valid is False:
                continue

            rhsNames = []
            rhsNumbers = []
            donor = None
            dateAdded = None
            metaData = None

            # Analyse the image file name to extract plant name, rhs number,
            # donor, date added and metadata
            print(f"* {imageNum+1}/{len(self.pendingPlantsImageInfo)}: '{imageInfo.filename}'")
            splitNames = imageInfo.filename.split('&&')
            for splitName in splitNames:
                name = None
                rhsNumber = 0
                imageData = re.search(r'(\D+(\(\S+\))?)\s(\d+)\s*(\D*)\s*(\d*)\s*(\D*)', splitName.strip())
                if imageData:
                    # Extract plant name, remove trailing spaces
                    name = imageData.group(1).strip()
                    # Extract RHS number
                    if len(imageData.groups()) > 2:
                        rhsNumber = int(imageData.group(3))
                    # Extract donor name
                    if len(imageData.groups()) > 3:
                        donor = imageData.group(4)
                    # Specify date added
                    if len(imageData.groups()) > 4 and imageData.group(5) and imageData.group(5) != '0':
                        dateAdded = f"01/01/{imageData.group(5)}"
                    # Extract meta data
                    if len(imageData.groups()) > 5:
                        metaData = imageData.group(6)

                # Get the RHS numbers of the image
                if name:
                    print(f"  - Got plant name extracted as '{name}'")
                    found = False
                    foundMatch = False
                    matchingNumbers = []
                    # See if we can find the name in the RHS database to give a best guess
                    for index in range(2, self.rhsReferenceDB.workbook['Table1'].max_row):
                        rhsName = self.rhsReferenceDB.getValue('Table1', index, self.RHS_CACLFULLNAME)
                        if rhsName:
                            if self.constainsName(name, rhsName):
                                found = True
                                matchingNumbers.append(self.rhsReferenceDB.getValue('Table1', index, self.RHS_OLDSPECIESCODE))
                                print(f"        -> found name in RHS dataset as number '{self.rhsReferenceDB.getValue('Table1', index, 1)}', name '{rhsName}'")
                    # We managed to extract an RHS number from the file name. Check
                    # if correct
                    if rhsNumber != 0:
                        print(f"    Got RHS number extracted as '{rhsNumber}'")
                        # See if we can find the number in the RHS database to give
                        # the expected name which we can compare with the name
                        # extracted from the file name
                        for index in range(2, self.rhsReferenceDB.workbook['Table1'].max_row):
                            oldspeciescode = self.rhsReferenceDB.getValue('Table1', index, self.RHS_OLDSPECIESCODE)
                            if oldspeciescode and int(oldspeciescode) == rhsNumber:
                                if rhsNumber in matchingNumbers:
                                    foundMatch = True
                                found = True
                                val = self.rhsReferenceDB.getValue('Table1', index, self.RHS_CACLFULLNAME)
                                print(f"        -> found number in the RHS dataset as RHS name  '{val}'")
                                break
                        if not found:
                            rhsNumber = 0
                        else:
                            if foundMatch:
                                rhsNumbers.append(rhsNumber)
                                rhsNames.append("")
                            else:
                                val = input(f"    Accept RHS number '{rhsNumber}'? (Y/n) ")
                                if not val or val.lower() == 'y':
                                    rhsNumbers.append(rhsNumber)
                                    rhsNames.append("")
                                else:
                                    rhsNumber = 0
                    # Either we didn't manage to extract a number from the file name
                    # or the number was wrong
                    if rhsNumber == 0:
                        # We didn't manage to extract an RHS number from the file name
                        val = input("    Specify RHS number ('enter' for unknown provenance ; for multiple, split by ','): ")
                        if not val:
                            print("    Put on list of unknown provenance'")
                            imageInfo.unknownProvenance = True
                            continue
                        numbers = val.split(',')
                        for num in numbers:
                            rhsNumbers.append(int(num))
                            rhsNames.append("")
                else:
                    print("    ! Couldn't extract name from file name. Need to rename file")
                    continue

            # Ignore images with plants of unknown provenance
            if imageInfo.unknownProvenance is True:
                continue

            # Check if we found any numbers
            if len(rhsNumbers) == 0:
                print("    ! Didn't find any valid RHS numbers. Ignoring image.")
                imageInfo.valid = False
                continue

            # Now that we have found some numbers, extract the information
            imageInfo.rhsNumbers = rhsNumbers
            rhsNumbersFound = 0
            for index, num in enumerate(rhsNumbers):
                # A value of '0' is valid, e.g. when there's no entry for
                # the plant in the RHS data set so no need to look anything up
                if num == 0:
                    imageInfo.rhsNames.append(rhsNames[index])
                    imageInfo.rhsHtml.append(self.createHtmlTag(rhsNames[index]))
                    rhsNumbersFound += 1  # technically not correct but makes it easier further down the line to pretend we did
                    continue
                # Find the data in the RHS data set for given RHS number
                for index in range(2, self.rhsReferenceDB.workbook['Table1'].max_row):
                    oldspeciescode = self.rhsReferenceDB.getValue('Table1', index, self.RHS_OLDSPECIESCODE)
                    if oldspeciescode and int(oldspeciescode) == num:
                        imageInfo.rhsFamily.append(self.rhsReferenceDB.getValue('Table1', index, self.RHS_FAMILYNAME))
                        imageInfo.rhsGenus.append(self.rhsReferenceDB.getValue('Table1', index, self.RHS_GENUSNAME))
                        imageInfo.rhsSpecies.append(self.rhsReferenceDB.getValue('Table1', index, self.RHS_SPECIESNAME))
                        imageInfo.rhsCultivar.append(self.rhsReferenceDB.getValue('Table1', index, self.RHS_CULTIVAR))
                        imageInfo.rhsNames.append(self.rhsReferenceDB.getValue('Table1', index, self.RHS_CACLFULLNAME))  # NAME
                        imageInfo.rhsHtml.append(self.createHtmlName(index))
                        rhsNumbersFound += 1
                        # Now that we have found the data, check if this is
                        # a new addition to the HPS library (interesting to
                        # know
                        samePlants = 0
                        for index in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
                            intnum = self.hpsPlantsDB.getValue('Plants', index, 3)  # RHS No
                            if not intnum:
                                continue
                            try:
                                int(intnum)
                            except ValueError:
                                continue
                            if int(intnum) == num:
                                samePlants += 1
                        if samePlants == 0:
                            newPlants += 1
                            print("  ! New plant in the list")
                        else:
                            print(f"    There are already {samePlants} images of this plant in the list")
                            # Now is the time to validate the image size as we
                            # want to add an image even if it's too small when
                            # there's no other in the list yet.
                            if imageInfo.validateSize():
                                print(f"  ! pending image {imageInfo.filename} is too small ({imageInfo.width}x{imageInfo.height})")
                                val = input("      Make invalid [YES/no] ? ")
                                if not val or val == 'YES' or val == 'yes':
                                    imageInfo.valid = False
                        break

            # Clearly something went wrong if we did't find all the numbers.
            # Can happen if the database is out of date and plant is on the RHS
            # website where we got this number from
            if rhsNumbersFound != len(rhsNumbers):
                if len(rhsNumbers) == 1:
                    print("        ! Can't find corresponding plant")
                else:
                    print("        ! Can't find all corresponding plants")
                val = input(f"  - Want to continue with {rhsNumbers} (y/N) ? ")
                if not val:
                    imageInfo.valid = False
                    continue

            # Check if donor name extracted from file name is correct
            if donor:
                val = input(f"  - Got donor as '{donor.rstrip()}'. Is this correct? (Y/n) ")
                # If a value is given (i.e. 'n') then delete donor name
                if val:
                    donor = None
            # If still no donor name then ask for it
            if not donor:
                donor = input("  - Please give donor name (note possible 'Anonymous' or 'Unknown'): ")
            imageInfo.donor = donor.rstrip()

            # If no date was extracted from the file name then take current date
            if not dateAdded:
                now = datetime.datetime.now()
                dateAdded = now.strftime("%d/%m/%Y")
            imageInfo.dateAdded = dateAdded
            print(f"  - Got date added as {dateAdded}")

            # Add any metadata
            if metaData:
                print(f"  - Got meta data added as '{metaData}'")
                imageInfo.metaData = metaData

        if newPlants > 0:
            print(f"! Got {newPlants} new plants")

        print()
        return 0

    def createAccession(self):
        if self.pendingPlantImages:
            maxRow = self.hpsPlantsDB.workbook['Plants'].max_row
            accession = int(self.hpsPlantsDB.getValue('Plants', maxRow, 2)[1:])  # Number
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid is False or imageInfo.unknownProvenance is True:
                    continue
                maxRow += 1
                accession += 1
                imageInfo.xlsxRow = maxRow
                imageInfo.accession = accession

        if self.pendingGardenImages:
            maxRow = self.hpsGardensDB.workbook['Gardens'].max_row
            accession = int(self.hpsGardensDB.getValue('Gardens', maxRow, 2)[1:])  # Number
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid is False:
                    continue
                maxRow += 1
                accession += 1
                imageInfo.xlsxRow = maxRow
                imageInfo.accession = accession

        return 0

    def updateSpreadsheets(self):
        print("Update spreadsheets")
        print("-------------------")

        if self.pendingPlantImages:
            backupHpsPlantsDB = self.hpsPlantsDB.filename+' - '+datetime.datetime.now().strftime("%d%m%y")+self.hpsPlantsDB.extension
            print(f"* Create backup of '{self.hpsPlantsDB.filename+self.hpsPlantsDB.extension}' to {backupHpsPlantsDB}")
            if os.path.exists(self.scriptDir+backupHpsPlantsDB):
                print("  - File already exists. Skipping")
            else:
                shutil.copyfile(self.scriptDir+self.hpsPlantsDB.filename+self.hpsPlantsDB.extension, self.scriptDir+backupHpsPlantsDB)
            print(f"* Update '{self.scriptDir+self.hpsPlantsDB.filename+self.hpsPlantsDB.extension}': master image database")
            validFiles = 0
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid is False or imageInfo.unknownProvenance is True:
                    continue
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 1, self.convertSpecialChar(imageInfo.getRHSName()))  # plantName
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 2, "P{:05d}".format(imageInfo.accession))  # HPSNo
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 3, imageInfo.getRHSNumber())  # RHSNo
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 8, imageInfo.donor)  # donor
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 9, imageInfo.dateAdded)  # dateAdded
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 11, imageInfo.metaData)  # extra information
                validFiles += 1
            if validFiles == 0:
                print("  ! No data to write")
            else:
                if not self.args.dryrun:
                    try:
                        self.hpsPlantsDB.save(self.scriptDir+self.hpsPlantsDB.filename+self.hpsPlantsDB.extension)
                    except PermissionError:
                        print(f"! Permission error writing to {self.scriptDir+self.hpsPlantsDB.filename+self.hpsPlantsDB.extension}")

            backupimagelibDB = self.imagelibDB.filename+' - '+datetime.datetime.now().strftime("%d%m%y")+self.imagelibDB.extension
            print(f"* Create backup of '{self.imagelibDB.filename+self.imagelibDB.extension}' to {backupimagelibDB}")
            if os.path.exists(self.scriptDir+backupimagelibDB):
                print("  - File already exists. Skipping")
            else:
                shutil.copyfile(self.scriptDir+self.imagelibDB.filename+self.imagelibDB.extension, self.scriptDir+backupimagelibDB)
            print(f"* Update '{self.scriptDir+self.imagelibDB.filename+self.imagelibDB.extension}': image database used by website")
            # Find where the P files (plants) change into X files (gardens)
            padd = 0
            for index in range(2, self.imagelibDB.workbook['active'].max_row):
                imageID = self.imagelibDB.getValue('active', index, 2)  # Image ID
                if imageID.startswith("X"):
                    padd = index
                    break
            validFiles = 0
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid is False or imageInfo.unknownProvenance is True:
                    continue
                # Insert plants at end of P files
                self.imagelibDB.workbook['active'].insert_rows(padd)
                html = "<span RHS>" + imageInfo.rhsHtml[0] + "</span>"
                for index in range(1, len(imageInfo.rhsHtml)):
                    html += " && <span RHS>" + imageInfo.rhsHtml[index] + "</span>"
                self.imagelibDB.setValue('active', padd, 1,  html)  # Caption
                self.imagelibDB.setValue('active', padd, 2, f"P{imageInfo.accession:05}")  # Image ID
                validFiles += 1
                padd += 1
            if validFiles == 0:
                print("! No data to write")
            else:
                if not self.args.dryrun:
                    try:
                        self.imagelibDB.save(self.scriptDir+self.imagelibDB.filename+self.imagelibDB.extension)
                    except PermissionError:
                        print("  ! Permission error writing to {}".format(self.imagelibDB.fileName))

            backupGeneraDB = self.generaDB.filename+' - '+datetime.datetime.now().strftime("%d%m%y")+self.generaDB.extension
            print(f"* Create backup of '{self.generaDB.filename+self.generaDB.extension}' to {backupGeneraDB}")
            if os.path.exists(self.scriptDir+backupGeneraDB):
                print("  - File already exists. Skipping")
            else:
                shutil.copyfile(self.scriptDir+self.generaDB.filename+self.generaDB.extension, self.scriptDir+backupGeneraDB)
            print(f"* Update '{self.scriptDir+self.generaDB.filename+self.generaDB.extension}': alphabetically sorted list of genera we have pictures of. Used by website.")
            existingGenus = self.generaDB.getColumn('active', 1)
            genusAdded = False
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid is False or imageInfo.unknownProvenance is True:
                    continue
                # Add the genus to the spreadsheet if it's not already there
                if len(imageInfo.rhsGenus) == 0:
                    continue
                for index in range(len(imageInfo.rhsGenus)):
                    genus = imageInfo.rhsGenus[index].capitalize()
                    if genus in existingGenus:
                        # genus already in database
                        continue
                    for rhsGenus in existingGenus:
                        # First 7 rows don't contain valid compare data
                        if existingGenus.index(rhsGenus)<8:
                            continue
                        if genus < rhsGenus:
                            print(f"  - Inserting new genus {genus}, family {imageInfo.rhsFamily[index].capitalize()} in row {existingGenus.index(rhsGenus)}")
                            self.generaDB.workbook['active'].insert_rows(existingGenus.index(rhsGenus)+1)
                            self.generaDB.setValue('active', existingGenus.index(rhsGenus)+1, 1, genus)
                            self.generaDB.setValue('active', existingGenus.index(rhsGenus)+1, 2, imageInfo.rhsFamily[index].capitalize())
                            genusAdded = True
                            # Recreate the list of existing genus as a new one has just been added
                            existingGenus = self.generaDB.getColumn('active', 1)
                            break
            if genusAdded:
                if not self.args.dryrun:
                    try:
                        self.generaDB.save(self.scriptDir+self.generaDB.filename+self.generaDB.extension)
                    except PermissionError:
                        print("  ! Permission error writing to {}".format(self.generaDB.fileName))
            else:
                print("  - No new genus added")

        if self.pendingGardenImages:
            backupHpsGardensDB = self.hpsGardensDB.filename+' - '+datetime.datetime.now().strftime("%d%m%y")+self.hpsGardensDB.extension
            print(f"* Create backup of '{self.hpsGardensDB.filename+self.hpsGardensDB.extension}' to {backupHpsGardensDB}")
            if os.path.exists(self.scriptDir+backupHpsGardensDB):
                print("  - File already exists. Skipping")
            else:
                shutil.copyfile(self.scriptDir+self.hpsGardensDB.filename+self.hpsGardensDB.extension, self.scriptDir+backupHpsGardensDB)
            print(f"* Update '{self.scriptDir+self.hpsGardensDB.filename+self.hpsGardensDB.extension}': master garden image database")
            validFiles = 0
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid is False:
                    continue
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 1, imageInfo.gardenName)  # plantName
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 2, "X{:05d}".format(imageInfo.accession))  # HPSNo
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 3, imageInfo.donor)  # donor
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 4, imageInfo.dateAdded)  # dateAdded
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 6, imageInfo.metaData)  # extra information
                validFiles += 1
            if validFiles == 0:
                print("  ! No data to write")
            else:
                if not self.args.dryrun:
                    try:
                        self.hpsGardensDB.save(self.scriptDir+self.hpsGardensDB.filename+self.hpsGardensDB.extension)
                    except PermissionError:
                        print(f"! Permission error writing to {self.scriptDir+self.hpsGardensDB.filename+self.hpsGardensDB.extension}")

            backupimagelibDB = self.imagelibDB.filename+' - '+datetime.datetime.now().strftime("%d%m%y")+self.imagelibDB.extension
            print(f"* Create backup of '{self.imagelibDB.filename+self.imagelibDB.extension}' to {backupimagelibDB}")
            if os.path.exists(self.scriptDir+backupimagelibDB):
                print("  - File already exists. Skipping")
            else:
                shutil.copyfile(self.scriptDir+self.imagelibDB.filename+self.imagelibDB.extension, self.scriptDir+backupimagelibDB)
            print(f"* Update '{self.scriptDir+self.imagelibDB.filename+self.imagelibDB.extension}': image database used by website")
            validFiles = 0
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid is False:
                    continue
                # Insert gardens at end of workbook
                row = self.imagelibDB.workbook['active'].max_row+1
                self.imagelibDB.setValue('active', row, 1, imageInfo.gardenName)  # Caption
                self.imagelibDB.setValue('active', row, 2, f"X{imageInfo.accession:05}")  # Image ID
                validFiles += 1
            if validFiles == 0:
                print("! No data to write")
            else:
                if not self.args.dryrun:
                    try:
                        self.imagelibDB.save(self.scriptDir+self.imagelibDB.filename+self.imagelibDB.extension)
                    except PermissionError:
                        print("  ! Permission error writing to {}".format(self.imagelibDB.fileName))

        print()
        return 0

    def copyImagesToUpload(self):
        print("Copy images")
        print("-----------")

        if self.pendingPlantImages:
            # Copy plant images into upload directory and remove GPS data
            print(f"* Copy plant images to '{self.uploadPlantsDir}'")
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid is False or imageInfo.unknownProvenance is True:
                    continue
                startletter = imageInfo.getRHSName()[0]
                # Copy from pending to dropbox upload directory
                if not self.args.dryrun:
                    os.makedirs(self.uploadPlantsDir+startletter, exist_ok=True)
                newFilename = self.uploadPlantsDir+startletter+"\\"+self.convertSpecialChar(imageInfo.getRHSName())+" P{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                if not self.args.dryrun:
                    try:
                        shutil.copy2(imageInfo.path, newFilename)
                    except OSError as e:
                        print(f"! Can't copy file. Error: {e}")
                        imageInfo.valid = False
                        continue
                    subprocess.Popen(["exiftool",
                                      "-gpsaltitude=",
                                      "-gpslatitude=",
                                      "-gpslongitude=",
                                      "-overwrite_original",
                                      newFilename],
                                     stdout=subprocess.PIPE).communicate()

        if self.pendingGardenImages:
            # Copy garden images into upload directory and remove GPS data
            print(f"* Copy garden images to upload directory '{self.uploadGardensDir}'")
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid is False:
                    continue
                # Copy from pending to dropbox upload directory
                if not self.args.dryrun:
                    os.makedirs(self.uploadGardensDir, exist_ok=True)
                newFilename = self.uploadGardensDir+"\\"+imageInfo.gardenName+" X{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                if not self.args.dryrun:
                    try:
                        shutil.copy2(imageInfo.path, newFilename)
                    except OSError as e:
                        print(f"! Can't copy file. Error: {e}")
                        imageInfo.valid = False
                        continue
                    subprocess.Popen(["exiftool",
                                      "-gpsaltitude=",
                                      "-gpslatitude=",
                                      "-gpslongitude=",
                                      "-overwrite_original",
                                      newFilename],
                                     stdout=subprocess.PIPE).communicate()

        # Create thumbnails: resize, auto orientate, remove exif, add watermark
        print(f"* Create thumbnails in {self.uploadThumbsDir}")
        # The watermark will appear in the middle bottom, white, offset by 12 pixels.
        watermarkText = "gravity south fill white text 0,12 'Hardy Plant Society\\nwww.hardy-plant.org.uk'"

        if not self.args.dryrun:
            os.makedirs(self.uploadThumbsDir, exist_ok=True)
        if self.pendingPlantImages:
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid is False or imageInfo.unknownProvenance is True:
                    continue
                startletter = imageInfo.getRHSName()[0]
                oldFilename = self.uploadPlantsDir+startletter+"\\"+self.convertSpecialChar(imageInfo.getRHSName())+" P{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                oldFilename = oldFilename.replace(u'/', u'_')
                newFilename = self.uploadThumbsDir+"P{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()

                # Convert image
                if not self.args.dryrun:
                    out = subprocess.Popen(["magick",
                                            oldFilename,
                                            "-resize", "350x350",   # Maximum size
                                            "-density", "72",       # DPI
                                            "-auto-orient",         # Orientation
                                            "-strip",               # Strip of any comments or profiles (e.g. exif)
                                            "-font", "Microsoft-Sans-Serif",
                                            "-pointsize", "8.25",
                                            "-draw", watermarkText,
                                            newFilename], stdout=subprocess.PIPE)
                    stdout, stderr = out.communicate()
                    if out.returncode != 0:
                        print("  ! Error")
                        imageInfo.valid = False
                        continue

            # Copy plant of unknown provenance into separate directory
            foundUnknownProvenance = False
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.unknownProvenance:
                    foundUnknownProvenance = True
                    break
            if foundUnknownProvenance:
                print(f"* Copy to upload directory for unknown provenance '{self.uploadUnknownProvenancePlantsDir}'")
                for imageInfo in self.pendingPlantsImageInfo:
                    if imageInfo.unknownProvenance is False:
                        continue
                    # Copy from pending to unknown provenance upload directory
                    if not self.args.dryrun:
                        os.makedirs(self.uploadUnknownProvenancePlantsDir, exist_ok=True)
                    newFilename = self.uploadUnknownProvenancePlantsDir+self.convertSpecialChar(imageInfo.filename)+imageInfo.extension
                    if not self.args.dryrun:
                        try:
                            shutil.copyfile(imageInfo.path, newFilename)
                        except OSError as e:
                            print(f"! Can't copy file. Error: {e}")
                            continue

        if self.pendingGardenImages:
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid is False:
                    continue
                oldFilename = self.uploadGardensDir+imageInfo.gardenName+" X{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                oldFilename = oldFilename.replace(u'/', u'_')
                newFilename = self.uploadThumbsDir+"X{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()

                # Convert image
                if not self.args.dryrun:
                    out = subprocess.Popen(["magick",
                                            oldFilename,
                                            "-resize", "350x350",   # Maximum size
                                            "-density", "72",       # DPI
                                            "-auto-orient",         # Orientation
                                            "-strip",               # Strip of any comments or profiles (e.g. exif)
                                            "-font", "Microsoft-Sans-Serif",
                                            "-pointsize", "8.25",
                                            "-draw", watermarkText,
                                            newFilename], stdout=subprocess.PIPE)
                    stdout, stderr = out.communicate()
                    if out.returncode != 0:
                        print("  ! Error")
                        imageInfo.valid = False
                        continue

        print()
        return 0

    def printFinalise(self):
        print("Finally")
        print("-------")
        print(f"Copy plant  images     to {self.plantsDir}")
        print(f"Copy garden images     to {self.gardensDir}")
        print(f"Copy all    thumbnails to {self.thumbsDir}")
        print()
        print("Upload plant  images     to Dropbox/Family Room/Image Library/Images - plants")
        print("Upload garden images     to Dropbox/Family Room/Image Library/Images - Gardens")
        print("Upload plant  thumbnails to Dropbox/Family Room/Image Library/Images - Plants - thumbs")
        print("Upload garden thumbnails to Dropbox/Family Room/Image Library/Images - Gardens - thumbs")
        print()
        print("Upload genera.csv               to Dropbox/Family Room/Image Library")
        print("       HPS Images - Gardens.xls")
        print("       HPS Images - Plants.xlsx")
        print("       imagelib.csv")
        print()
        print("Email marketing@hardy-plant.org.uk to let them know you've added new images.")


################################################################################


def main():
    # Process the arguments
    parser = argparse.ArgumentParser(
        description='Process new HPS images.',
        formatter_class=argparse.RawTextHelpFormatter,
        epilog='''
Usage
-----
The base directory where all images can be found from is defined in self.baseDir

It can then find the directories containing all the plant images, garden images,
pending plant images, pending garden images and the upload directory from that
base directory.

The script will go through a number of steps to classify images:
* The first step is checking if the tools the script needs have been downloaded
  ('magick' and 'exiftool'). If it can't find it then it will tell where to
  download it from.
* The second step is validating the input: are all the directories it expects
  there, are all the spreadsheets it expects there and are they in the format it
  needs.
* It will import all the existing and pending images in the next step

Image file names are in the form of

    <plant name> <RHS number> <donor> <year> <misc>.jpg

If multiple plants, the form is

    <plant name> <RHS number> <donor> <year> && <plant name> <RHS number> <donor> <year> <misc>.jpg

If 'RHS number' not known, use '0'
'donor' can be 'unknown' or 'anonymous'
If 'year' not specified or 0 then the current date will be used instead.

If you have two different pictures of the same plant of the same donor then I
would suggest adding an extra space before/after the RHS number
''')
    parser.add_argument(
        '--dryrun',
        action='store_true',
        help='Run without saving/creating any files'
    )
    args = parser.parse_args()

    # Construct the base class
    hps = CHPS(args)

    print()

    # Check if required tools exist
    if hps.validateTools():
        return 1

    # Check if required directories and xlsx files exist and are valid
    if hps.validateInput():
        return 1

    # Import the pending and HPS library images
    if hps.importImages():
        return 1

    # Get the RHS numbers of the pending images
    if hps.updateImageInfo():
        return 1

    # Create accession numbers for all the pending images
    if hps.createAccession():
        return 1

    # Copy the pending images to new directory, ready to be uploaded
    if hps.copyImagesToUpload():
        return 1

    if hps.updateSpreadsheets():
        return 1

    hps.printFinalise()

    return 0


if __name__ == "__main__":
    ret = main()
    sys.exit(ret)

sys.exit(1)
