#!/usr/bin/python
from CImageInfo   import CImageInfo
from CImageInfo   import CPendingImageInfo
from CSpreadSheet import CSpreadSheet

import argparse
import datetime
import os
import re
import shutil
import subprocess
import sys
import urllib.request

# Force print to always flush
import functools
print = functools.partial(print, flush=True)

class CHPS:
    def __init__(self, args):
        self.args = args

        self.gitHubDir         = 'H:\\hps_categorise\\'
        self.baseDir           = 'H:\\HPS_Images\\'
        self.plantsDir         = self.baseDir + 'Plants\\'
        self.pendingPlantsDir  = self.baseDir + 'Pending\\Plants\\'
        self.gardensDir        = self.baseDir + 'Gardens\\'
        self.pendingGardensDir = self.baseDir + 'Pending\\Gardens\\'
        self.thumbsDir         = self.baseDir + 'Thumbnails\\'
        self.uploadDir         = self.baseDir + 'Upload\\'

        # Current data
        self.hpsPlantsImageInfo  = []
        self.hpsGardensImageInfo = []
        # Pending data
        self.pendingPlantImages      = True
        self.pendingPlantsImageInfo  = []
        self.pendingGardenImages     = True
        self.pendingGardensImageInfo = []
        # Upload directories
        now = datetime.datetime.now()
        self.uploadDropBoxDir        = self.uploadDir+'To_DropBox_'+now.strftime("%d%m%y")+'\\'
        self.uploadDropBoxPlantsDir  = self.uploadDropBoxDir+'Plants\\'
        self.uploadDropBoxGardensDir = self.uploadDropBoxDir+'Gardens\\'
        self.uploadFtpDir            = self.uploadDir+'To_FTP_'+now.strftime("%d%m%y")+'\\'
        self.uploadFtpThumbsDir      = self.uploadFtpDir+'thumbs\\'
        self.uploadUnknownProvenancePlantsDir = self.uploadDir+'unknownProvenance_'+now.strftime("%d%m%y")+'\\'

    def stats(self, startCount):
        print(f"Analysis")
        print( "--------")

        # Get the RHS database
        if self.createRhsReferenceDB() == 0:
            rhsNumbers = set()
            rhsGenera  = set()
            for index in range(2, self.rhsReferenceDB.workbook['HPS-NAMES May 19'].max_row):
                rhsNumber = self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 1)
                if rhsNumber:
                    rhsNumbers.add(rhsNumber)
                rhsGenus  = self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 7)
                if rhsGenus:
                    rhsGenera.add(rhsGenus)
            print( "  * Overall in RHS library, there are:")
            print(f"    - {len(rhsGenera)} genus")
            print(f"    - {len(rhsNumbers)} species")
            print()

        # get the HPS database
        if self.createHpsPlantsDB():
            return 1

        totalNumImages = 0
        numPlants = 0
        numNewImages = 0
        numNewPlants = 0
        foundStartCount = False
        rhsNumbers = set()
        hpsDonors = set()
        hpsGenus = set()
        addedGenus = set()
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            # Ignore the withdrawn images
            RHSNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 3) # RHS no
            if not RHSNumber: continue
            if RHSNumber == 'WITHDRAWN': continue
            dateWithdrawn = self.hpsPlantsDB.getValue('Plants', currentRow, 12) # Date withdrawn
            if dateWithdrawn: continue

            HPSName   = self.hpsPlantsDB.getValue('Plants', currentRow, 1) # HPS name
            HPSNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 2) # HPS no
            donor     = self.hpsPlantsDB.getValue('Plants', currentRow, 8) # Donor
            genus     = re.search(r'(\S+)', HPSName)
            if not genus: continue

            totalNumImages += 1

            # Find the point from where we want to start counting
            if HPSNumber == startCount:
                foundStartCount = True

            # Continue if we haven't found the start yet
            if not foundStartCount:
                rhsNumbers.add(RHSNumber)
                hpsGenus.add(genus.group(1))
                continue

            # Find number of totally new plants added since start count
            if RHSNumber not in rhsNumbers:
                numNewPlants += 1

            # Find the donors since start count
            if donor != 'Anonymous' and donor != 'Unknown':
                hpsDonors.add(donor)

            # Add the genus
            if genus.group(1) not in hpsGenus:
                addedGenus.add(genus.group(1))

            hpsGenus.add(genus.group(1))
            rhsNumbers.add(RHSNumber)
            numNewImages +=1

        print()
        print( "  * Overall in HPS library, there are:")
        print(f"    - {len(hpsGenus)} different genus")
        print(f"    - {len(rhsNumbers)} different species")
        print(f"    - {totalNumImages} valid images")
        print()
        print(f"  * Since {startCount}:")
        print(f"    - {numNewPlants}/{numNewPlants/(totalNumImages-numNewPlants)*100.0:.1f}% new taxa have been added not previously in library")
        print(f"    - {len(hpsDonors)} donors contributed")
        if len(addedGenus):
            print(f"    - {len(addedGenus)}/{len(addedGenus)/(len(hpsGenus)-len(addedGenus))*100.0:.1f}% new genera have been added not previously in library: {addedGenus}")
        else:
            print("    - 0 new genus have been added not previously in library")
        print(f"    - {numNewImages}/{numNewImages/(totalNumImages-numNewImages)*100.0:.1f}% images have been added")
        return 0

    def fullAnalysis(self):
        print("Analyse databases")
        print("-----------------")

        # Validate HPS plants
        if self.createHpsPlantsDB():
            return 1

        # Simple validation
        expectedHeaders = ["Plant name",    "Number",            "RHS no",
                           "RHS status",    "qualifier",         "descriptor",
                           "image caption", "Donor",             "Date added",
                           "Slide No.",     "Extra information", "Date withdrawn"]
        if self.hpsPlantsDB.validate('Plants', expectedHeaders):
            return 1

        # Check if plant name is filled in for each row
        print("    - Check each row for missing plant name")
        missingNameRows = []
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            name = self.hpsPlantsDB.getValue('Plants', currentRow, 1) # Plant name
            if not name:
                missingNameRows.append(currentRow)
        if len(missingNameRows):
            print(f"        ! Warning ! '{self.hpsPlantsDB.filename}' has {len(missingNameRows)} rows with missing plant names: {missingNameRows}")
        else:
            print("        No rows have missing plant names")

        # Check if image number is filled in
        print(f"    - Check each row for missing image numbers")
        missingImageNumberRows = []
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            imageNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 2) # Number
            if not imageNumber:
                missingImageNumberRows.append(currentRow)
        if len(missingImageNumberRows):
            print(f"        ! Warning! '{self.hpsPlantsDB.filename}' has {len(missingImageNumberRows)} rows with missing image numbers: {missingImageNumbersRows}")
        else:
            print("        No rows have missing image numbers")

        # Check if image number is valid
        print("    - Check each row for valid HPS image numbers")
        invalidImageNumberRows = []
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            imageNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 2) # Number
            if not imageNumber: continue
            elif not re.search(r'^(P|X)\d{5}$', imageNumber):
                invalidImageNumberRows.append(currentRow)
        if len(invalidImageNumberRows):
            print(f"        ! Warning ! '{self.hpsPlantsDB.filename}' has {len(invalidImageNumberRows)} rows with invalid image numbers: {invalidImageNumberRows}")
        else:
            print("        All rows have valid image numbers")

        # Check if RHS number is filled in and if not that there's at least
        # information as to why not
        print("    - Check each row with missing RHS numbers for given reason")
        missingRHSNumberRows = []
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            RHSNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 3) # RHS no
            if not RHSNumber:
                extraInformation = self.hpsPlantsDB.getValue('Plants', currentRow, 11) # Extra information
                if not extraInformation:
                    missingRHSNumberRows.append(currentRow)
        if len(missingRHSNumberRows):
            print(f"        ! Warning ! {len(missingRHSNumberRows)} rows have missing RHS numbers without reason given: {missingRHSNumberRows}")
        else:
            print("        All rows with missing RHS numbers have a reason")

        # Check if rhs number is set to withdrawn and has a withdrawn date
        print("    - Check if withdrawn notifications are valid")
        mismatchWithdrawnRows = []
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            if currentRow in missingRHSNumberRows: continue
            RHSNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 3) # RHS no
            dateWithdrawn = self.hpsPlantsDB.getValue('Plants', currentRow, 12) # Date withdrawn
            if RHSNumber != 'WITHDRAWN' and dateWithdrawn:
                mismatchWithdrawnRows.append(currentRow)
            if RHSNumber == 'WITHDRAWN' and not dateWithdrawn:
                mismatchWithdrawnRows.append(currentRow)
        if len(mismatchWithdrawnRows):
            print(f"        ! Warning ! HPS plants has {len(mismatchWithdrawnRows)} rows with mismatching withdrawn notifications: {mismatchWithdrawnRows}")
        else:
            print("        All withdrawn notifications are valid")

        # Check withdrawn files don't exists
        print("    - Make sure withdrawn image files have been removed")
        firstMissing = True
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            RHSNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 3) # RHS no
            if not RHSNumber: continue
            name = self.hpsPlantsDB.getValue('Plants', currentRow, 1) # Plant name
            if not name: continue
            imageNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 2) # Number
            if not imageNumber: continue
            dateWithdrawn = self.hpsPlantsDB.getValue('Plants', currentRow, 12) # Date withdrawn
            if RHSNumber == 'WITHDRAWN' or dateWithdrawn:
                fileName = self.plantsDir + name[0] + "\\" + name.replace('/','_') + " " + imageNumber + ".jpg"
                if os.path.isfile(fileName):
                    if firstMissing:
                        firstMissing=False
                    print(f"        Found withdrawn file '{fileName}'")
        if firstMissing:
            print("        All withdrawn image files have been removed")

        # Check if file exists
        print("    - Check if all valid image files exist")
        missingFileRows = []
        firstMissing = True
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            RHSNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 3) # RHS no
            if not RHSNumber: continue
            name = self.hpsPlantsDB.getValue('Plants', currentRow, 1) # Plant name
            if not name: continue
            imageNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 2) # Number
            if not imageNumber: continue
            dateWithdrawn = self.hpsPlantsDB.getValue('Plants', currentRow, 12) # Date withdrawn
            if RHSNumber == 'WITHDRAWN' or dateWithdrawn:
                continue
            fileName = self.plantsDir
            if name.startswith("x "):
                fileName += name[2]
            else:
                fileName += name[0]
            fileName += "\\" + name.replace('/','_') + " " + imageNumber + ".jpg"
            if not os.path.isfile(fileName):
                if firstMissing:
                    firstMissing = False
                print(f"        Can't find '{fileName}'")
                missingFileRows.append(currentRow)
        if len(missingFileRows) == 0:
            print("        All valid image files exist")
        print()

        # Validate RHS dataset
        if self.createRhsReferenceDB():
            return 1

        # Simple validation
        expectedHeaders = ["NAME_NUM", "ACCEPT_FULL", "NAME",
                           "AWARD", "ALT_NAME_FULL", "FAMILY",
                           "GENUS", "GEN_HYBR", "SPECIES",
                           "SPEC_AUTH", "SPEC_HYBR", "INFRA_RANK_FULL",
                           "INFRA_EPI", "INFRA_AUTH", "CULTIVAR",
                           "CULTIVAR_AUTH", "CV_FLAG", "CV_GROUP",
                           "SOLD_AS", "DESCRIPTOR", "IDENT_QUAL_FULL",
                           "AGG_FLAG_FULL", "GENUS_2", "SPECIES_2",
                           "SPEC_AUTH_2", "INFRA_RANK_2_FULL", "INFRA_EPI_2",
                           "INFRA_AUTH_2", "CULTIVAR_2", "CULTIVAR_AUTH_2",
                           "CV_FLAG_2", "CV_GROUP_2", "SOLD_AS_2",
                           "DESCRIPTOR_2", "IDENT_QUAL_FULL_2", "AGG_FLAG_FULL_2",
                           "NAME_FREE", "GROUP_NAME", "GROUP_NAME_FULL",
                           "PARENTAGE", "ALT_NAME", "NAME_HTML",
                           "USER3"]
        if self.rhsReferenceDB.validate('HPS-NAMES May 19', expectedHeaders):
            return 1
        print()

        # Cross reference RHS numbers in HPS Plants with the name in RHS Reference
        # and plant name in HPS plants
        print(f"  - Cross reference RHS numbers in '{self.hpsPlantsDB.filename}' with RHS names in '{self.rhsReferenceDB.filename}'")
        wrongNumbers = []
        wrongNames   = []
        numbers    = self.hpsPlantsDB.getColumn('Plants', 3) # RHS no
        hpsNames   = self.hpsPlantsDB.getColumn('Plants', 1) # Plant name
        rhsNumbers = self.rhsReferenceDB.getColumn('HPS-NAMES May 19', 1) # NAME_NUM
        rhsNames   = self.rhsReferenceDB.getColumn('HPS-NAMES May 19', 3) # NAME
        for hpsIndex in range(1, len(numbers)):
            imageNumbers = numbers[hpsIndex]
            if not imageNumbers:
                continue
            if "&&" in str(imageNumbers):
                numberList = re.findall(r'\d{1,6}', imageNumbers)
            else:
                numberList = [ str(imageNumbers) ]
            for imageNumber in numberList:
                if imageNumber and imageNumber.isnumeric():
                    if int(imageNumber) not in rhsNumbers:
                        #print(f"  RHS number '{imageNumber}' doesn't exist in RHS list")
                        wrongNumbers.append(hpsIndex+1)
                    else:
                        rhsIndex = rhsNumbers.index(int(imageNumber))
                        rhsName = rhsNames[rhsIndex]
                        rhsName = re.sub('\s*AGM', '', rhsName)
                        rhsName = re.sub('\s*\(PBR\)', '', rhsName)
                        if rhsName not in hpsNames[hpsIndex]:
                            wrongNames.append(hpsIndex+1)
                            #print('{}: "{}", "{}"'.format(hpsIndex, rhsName, hpsNames[hpsIndex]))
        if len(wrongNumbers):
            print(f"      ! Warning ! {len(wrongNumbers)} rows with invalid RHS numbers in '{self.hpsPlantsDB.filename}': {wrongNumbers}")
        if len(wrongNames):
            print(f"      !  Warning ! {len(wrongNames)} rows with invalid names in '{self.hpsPlantsDB.filename}': {wrongNames}")
        print()

        # Validate Imagelib
        if self.createImagelibDB():
            return 1
        expectedHeaders = ["Caption",
                           "Image ID"]
        if self.imagelibDB.validate('active', expectedHeaders):
            return 1
        print()

        print(f"  - Cross reference if withdrawn images in '{self.hpsPlantsDB.filename}' aren't in '{self.imagelibDB.filename}'")
        imagelibIDs = self.imagelibDB.getColumn('active', 2) # Image ID
        extraNumbers = []
        for currentRow in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
            imageNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 2) # Number
            if not imageNumber: continue
            # Make sure no withdrawn numbers are in imagelib
            RHSNumber = self.hpsPlantsDB.getValue('Plants', currentRow, 3) # RHS no
            if RHSNumber == 'WITHDRAWN' and imageNumber in imagelibIDs:
                extraNumbers.append(currentRow)
                continue
            dateWithdrawn = self.hpsPlantsDB.getValue('Plants', currentRow, 12) # Date withdrawn
            if dateWithdrawn and imageNumber in imagelibIDs:
                extraNumbers.append(currentRow)
                continue
        if len(extraNumbers):
            print(f"      ! Warning ! '{self.imagelibDB.filename}' has {len(extraNumbers)} entries which were withdrawn: {extraNumbers}")
        else:
            print(f"      No withdrawn images were found in '{self.imagelibDB.filename}'")
        print()

        print(f"  - Cross reference if HTML plant names in '{self.imagelibDB.filename}' match up with plant names in '{self.rhsReferenceDB.filename}'")
        for currentRow in range(2, self.imagelibDB.workbook['active'].max_row):
            imagelibName = self.imagelibDB.getValue('active', currentRow, 1) # Caption
            imagelibNumber = self.imagelibDB.getValue('active', currentRow, 2) # Image ID
            foundName = False
            # Find image number in HPS images
            for plantindex in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
                hpsPlantsNumber = self.hpsPlantsDB.getValue('Plants', plantindex, 2) # HPS Number
                if hpsPlantsNumber == imagelibNumber:
                    # Find RHS number in HPS images
                    hpsPlantsRHSNumber = self.hpsPlantsDB.getValue('Plants', plantindex, 3) # RHS Number
                    # Find RHS name in RHS database
                    for rhsindex in range(2, self.rhsReferenceDB.workbook['HPS-NAMES May 19'].max_row):
                        rhsNumber = self.rhsReferenceDB.getValue('HPS-NAMES May 19', rhsindex, 1)
                        if rhsNumber == hpsPlantsRHSNumber:
                            foundName = True
                            rhsHTMLName = "<span RHS>" + self.rhsReferenceDB.getValue('HPS-NAMES May 19', rhsindex, 42) + "</span>"
                            if rhsHTMLName != imagelibName:
                                print(f"      Names don't correspond for HPS image ID {imagelibNumber}, RHS number {rhsNumber}:")
                                print(f"          RHS name: {rhsHTMLName}")
                                print(f"          HPS name: {imagelibName}")
                        if foundName: break
                if foundName: break

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

        # TODO: Check if Unclassified directory exists

        return 0

    def createImagelibDB(self):
        # imagelib.csv can be found in docsftp@hardy-plant.org.uk:/plants
        fileName = self.gitHubDir+"imagelib.csv"
        print(f"  - {fileName} (needs to be downloaded manually): ", end="")
        if not os.path.isfile(fileName):
            print("doesn't exist!")
            return 1

        self.imagelibDB = CSpreadSheet(fileName)
        print("OK")

        return 0

    def createGeneraDB(self):
        # genera.csv can be found in docsftp@hardy-plant.org.uk:/plants
        fileName = self.gitHubDir+"genera.csv"
        print(f"  - {fileName} (needs to be downloaded manually): ", end="")
        if not os.path.isfile(fileName):
            print("doesn't exist!")
            return 1

        self.generaDB = CSpreadSheet(fileName)
        print("OK")

        return 0

    def createHpsPlantsDB(self):
        # Get the latest version of the plants database. This needs to be done
        # better by checking if files are same or not using requests
        fileName = self.gitHubDir+"HPS Images - Plants.xlsx"
        if self.args.download:
            print(f"  - {fileName}: downloading", end="\r")
            url = 'https://www.dropbox.com/scl/fi/g9x3a92tzociye8r0ors4/HPS%20Images%20-%20plants.xlsx?dl=1'
            #url = 'https://www.dropbox.com/sh/xu2j3xd9kyhlsbt/AADvj2H4gihukQfAXt6ymvtya/Digital%20Filing%20List%20MASTER.xlsx?dl=1'
            with urllib.request.urlopen(url) as f:
                data = f.read()
            try:
                with open(fileName, "wb") as f:
                    f.write(data)
            except PermissionError:
                print(f"  - {fileName}: don't have write permission. File already open?")
                return 1

        print(f"  - {fileName}: importing  ", end="\r")
        self.hpsPlantsDB = CSpreadSheet(fileName)
        print(f"  - {fileName}: OK         ")

        return 0

    def createHpsGardensDB(self):
        # Get the latest version of the gardens database. This needs to be done
        # better by checking if files are same or not using requests
        fileName = self.gitHubDir+"HPS Images - Gardens.xlsx"
        if self.args.download:
            print(f"  - {fileName}: downloading", end="\r")
            url = 'https://www.dropbox.com/scl/fi/c3sfnvn582oh24juluag3/HPS%20Images%20-%20gardens.xlsx?dl=1'
            #url = 'https://www.dropbox.com/sh/pe5a97q1m0p525q/AABjfLfHbxCtYiN_wlGw2qcda/Digital%20filing%20list%20MASTER%20-%20scenes.xlsx?dl=1'
            f = urllib.request.urlopen(url)
            data = f.read()
            f.close()
            try:
                with open(fileName, "wb") as f:
                    f.write(data)
            except PermissionError:
                print(f"  - {fileName}: don't have write permission. File already open?")
                return 1

        print(f"  - {fileName}: importing  ", end="\r")
        self.hpsGardensDB = CSpreadSheet(fileName)
        print(f"  - {fileName}: OK         ")

        return 0

    def checkConsistency(self):
        # Find where the P files (plants) change into X files (gardens)
        imagelibAccession = ""
        for index in range(2, self.imagelibDB.workbook['active'].max_row):
            imageID = self.imagelibDB.getValue('active', index, 2) # Image ID
            if imageID.startswith("X"):
                break
            imagelibAccession = imageID

        if self.pendingPlantImages:
            maxRow = self.hpsPlantsDB.workbook['Plants'].max_row
            hpsPlantsAccession = self.hpsPlantsDB.getValue('Plants', maxRow, 2) # Number

            if imagelibAccession != hpsPlantsAccession:
                print(f"* Libraries not consistent! {imagelibAccession} vs {hpsPlantsAccession}")
                return 1

        if self.pendingGardenImages:
            imagelibAccession = self.imagelibDB.getValue('active', self.imagelibDB.workbook['active'].max_row, 2) # Image ID
            maxRow = self.hpsGardensDB.workbook['Gardens'].max_row
            hpsGardensAccession = self.hpsGardensDB.getValue('Gardens', maxRow, 2) # Number

            if imagelibAccession != hpsGardensAccession:
                print(f"* Libraries not consistent! {imagelibAccession} vs {hpsGardensAccession}")
                return 1

        return 0

    def createRhsReferenceDB(self):
        # Get the latest version of the RHS database. This needs to be done better
        # by checking if files are same or not using requests
        fileName = self.gitHubDir+"RHS_Dataset.xlsx"
        if self.args.download:
            print(f"  - {fileName}: downloading", end="\r")
            url = 'https://www.dropbox.com/s/9n9cjd1ru27jjma/HPS-NAMES%20May%2019.xlsx?dl=1'
            f = urllib.request.urlopen(url)
            data = f.read()
            f.close()
            try:
                with open(fileName, "wb") as f:
                    f.write(data)
            except PermissionError:
                print(f"  - {fileName}: don't have write permission. File already open?")
                return 1

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
            expectedHeaders = ["NAME_NUM", "ACCEPT_FULL", "NAME",
                               "AWARD", "ALT_NAME_FULL", "FAMILY",
                               "GENUS", "GEN_HYBR", "SPECIES",
                               "SPEC_AUTH", "SPEC_HYBR", "INFRA_RANK_FULL",
                               "INFRA_EPI", "INFRA_AUTH", "CULTIVAR",
                               "CULTIVAR_AUTH", "CV_FLAG", "CV_GROUP",
                               "SOLD_AS", "DESCRIPTOR", "IDENT_QUAL_FULL",
                               "AGG_FLAG_FULL", "GENUS_2", "SPECIES_2",
                               "SPEC_AUTH_2", "INFRA_RANK_2_FULL", "INFRA_EPI_2",
                               "INFRA_AUTH_2", "CULTIVAR_2", "CULTIVAR_AUTH_2",
                               "CV_FLAG_2", "CV_GROUP_2", "SOLD_AS_2",
                               "DESCRIPTOR_2", "IDENT_QUAL_FULL_2", "AGG_FLAG_FULL_2",
                               "NAME_FREE", "GROUP_NAME", "GROUP_NAME_FULL",
                               "PARENTAGE", "ALT_NAME", "NAME_HTML",
                               "USER3"]
            if self.rhsReferenceDB.validate('HPS-NAMES May 19', expectedHeaders):
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
            print(f"  - imported current plant images{' ': <108}")

        if self.pendingGardenImages:
            for filename in os.listdir(self.gardensDir):
                fullpath =self.gardensDir + '\\' + filename
                self.hpsGardensImageInfo.append(CImageInfo(fullpath, False))
            print(f"  - imported current garden images{' ': <108}")

    def importPendingImages(self):
        if self.pendingPlantImages:
            for filename in os.listdir(self.pendingPlantsDir):
                fullpath = self.pendingPlantsDir + filename
                print(f"  - {filename: <108}", end="\r")
                self.pendingPlantsImageInfo.append(CPendingImageInfo(fullpath))
            print(f"  - imported pending plant images{' ': <108}")
        if self.pendingGardenImages:
            for filename in os.listdir(self.pendingGardensDir):
                fullpath = self.pendingGardensDir + filename
                print(f"  - {filename: <108}", end="\r")
                self.pendingGardensImageInfo.append(CPendingImageInfo(fullpath))
            print(f"  - imported pending garden images{' ': <108}")

    def getImageInfo(self):
        print("* Import existing images")
        self.importCurrentImages()
        if self.pendingPlantImages and len(self.hpsPlantsImageInfo)==0:
            print("! Couldn't find information for current plant images\n")
            return 1
        if self.pendingGardenImages and len(self.hpsGardensImageInfo)==0:
            print("! Couldn't find information for current garden images\n")
            return 1

        print("* Import pending images")
        self.importPendingImages()
        if self.pendingPlantImages and len(self.pendingPlantsImageInfo)==0:
            print("! Couldn't find information for pending plant images\n")
            return 1
        if self.pendingGardenImages and len(self.pendingGardensImageInfo)==0:
            print("! Couldn't find information for pending garden images\n")
            return 1

        return 0

    def importImages(self):
        print("Import images")
        print("-------------")

        # Get file information for pending and existing files
        if self.getImageInfo():
            return 1

        # Find unique pending plant pictures.
        # This has currently been switched off as too slow. Can't compare sizes
        # as stored images may have some exif data removed and therefore have
        # different size even if the image itself is the same
        print("* Check if pending plants are unique")
        print("  ! Switched off for now")
        #uniqueFiles = 0
        #for pendingImage in self.pendingPlantsImageInfo:
        #    print(f"  - '{pendingImage.filename}'{' ': <108}", end="\r")
        #    isUnique = True
        #    for currentImage in self.hpsPlantsImageInfo:
        #        if pendingImage.getSize() == currentImage.getSize():
        #            # Check if hash is same. If it is, go to next pending file
        #            if pendingImage.getFileHash() == currentImage.getFileHash():
        #                print(f"  ! Found same image: '{currentImage.filename+currentImage.extension}'{' ': <108}")
        #                pendingImage.valid = False
        #                isUnique = False
        #                break
        #    if isUnique:
        #        uniqueFiles += 1
        #if uniqueFiles==0:
        #    print("! Did't find any valid pending unique files")
        #    return 1

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

        value = value.replace(u'\N{LATIN SMALL LETTER A WITH DIAERESIS}',  u'a') # E4
        value = value.replace(u'\N{LATIN SMALL LETTER E WITH GRAVE}',      u'e') # E8
        value = value.replace(u'\N{LATIN SMALL LETTER E WITH ACUTE}',      u'e') # E9
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
            if imageInfo.valid == False:
                continue

            print(f"  - {imageNum+1}/{len(self.pendingGardensImageInfo)}: '{imageInfo.filename}'")
            name      = None
            imageData = re.search(r'(\D+)\s+\d+\s+(\D+)\s*(\d*)', imageInfo.filename.strip())
            if imageData:
                # Extract garden name
                imageInfo.gardenName = imageData.group(1)
                # Extract donor name
                imageInfo.donor = imageData.group(2).rstrip()
                # Extract date added
                if len(imageData.groups())>2 and imageData.group(3)!='0':
                    dateAdded = f"01/01/{imageData.group(3)}"
                # Extract information
                if len(imageData.groups())>3:
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
            if imageInfo.valid == False:
                continue;

            rhsNames   = []
            rhsNumbers = []
            donor      = None
            dateAdded  = None
            metaData   = None

            # Analyse the image file name to extract plant name, rhs number,
            # donor, date added and metadata
            print(f"* {imageNum+1}/{len(self.pendingPlantsImageInfo)}: '{imageInfo.filename}'")
            splitNames = imageInfo.filename.split('&&')
            for splitName in splitNames:
                name      = None
                rhsNumber = 0
                imageData = re.search(r'(\D+(\(\S+\))?)\s(\d+)\s*(\D*)\s*(\d*)\s*(\D*)', splitName.strip())
                if imageData:
                    # Extract plant name
                    name = imageData.group(1)
                    # Extract RHS number
                    if len(imageData.groups())>2:
                        rhsNumber = int(imageData.group(3))
                    # Extract donor name
                    if len(imageData.groups())>3:
                        donor = imageData.group(4)
                    # Specify date added
                    if len(imageData.groups())>4 and imageData.group(5)!='0':
                        dateAdded = f"01/01/{imageData.group(5)}"
                    # Extract meta data
                    if len(imageData.groups())>5:
                        metaData = imageData.group(6)

                # Get the RHS numbers of the image
                if name:
                    print(f"  - Got plant name extracted as '{name}'")
                    found = False
                    foundMatch = False
                    matchingNumbers = []
                    # See if we can find the name in the RHS database to give a best guess
                    for index in range(2, self.rhsReferenceDB.workbook['HPS-NAMES May 19'].max_row):
                        rhsName = self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 3)
                        if rhsName:
                            if self.constainsName(name, rhsName):
                                found = True
                                matchingNumbers.append(self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 1))
                                print(f"        -> found name in RHS dataset as number '{self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 1)}', name '{rhsName}'")
                    # We managed to extract an RHS number from the file name. Check
                    # if correct
                    if rhsNumber != 0:
                        print(f"    Got RHS number extracted as '{rhsNumber}'")
                        # See if we can find the number in the RHS database to give
                        # the expected name which we can compare with the name
                        # extracted from the file name
                        for index in range(2, self.rhsReferenceDB.workbook['HPS-NAMES May 19'].max_row):
                            if int(self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 1)) == rhsNumber:
                                if rhsNumber in matchingNumbers:
                                    foundMatch = True
                                found = True
                                val = self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 3)
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
                                if not val or val == 'Y':
                                    rhsNumbers.append(rhsNumber)
                                    rhsNames.append("")
                                else:
                                    rhsNumber = 0
                    # Either we didn't manage to extract a number from the file name
                    # or the number was wrong
                    if rhsNumber == 0:
                        # We didn't manage to extract an RHS number from the file name
                        val = input(f"    Specify RHS number ('enter' for unknown provenance ; for multiple, split by ','): ")
                        if not val:
                            print(f"    Put on list of unknown provenance'")
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
            if imageInfo.unknownProvenance == True:
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
                if num==0:
                    imageInfo.rhsNames.append(rhsNames[index])
                    imageInfo.rhsHtml.append(self.createHtmlTag(rhsNames[index]))
                    rhsNumbersFound += 1 # technically not correct but makes it easier further down the line to pretend we did
                    continue
                # Find the data in the RHS data set for given RHS number
                for index in range(2, self.rhsReferenceDB.workbook['HPS-NAMES May 19'].max_row):
                    if int(self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 1)) == num:
                        imageInfo.rhsFamily.append(  self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 6))  # FAMILY
                        imageInfo.rhsGenus.append(   self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 7))  # GENUS
                        imageInfo.rhsSpecies.append( self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 9))  # SPECIES
                        imageInfo.rhsCultivar.append(self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 15)) # CULTIVAR
                        imageInfo.rhsStatus.append(  self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 2))  # ACCEPT_FULL
                        imageInfo.rhsNames.append(   self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 3))  # NAME
                        imageInfo.rhsHtml.append(    self.rhsReferenceDB.getValue('HPS-NAMES May 19', index, 42)) # NAME_HTML
                        rhsNumbersFound += 1
                        # Now that we have found the data, check if this is
                        # a new addition to the HPS library (interesting to
                        # know
                        samePlants = 0
                        for index in range(2, self.hpsPlantsDB.workbook['Plants'].max_row):
                            intnum = self.hpsPlantsDB.getValue('Plants', index, 3) # RHS No
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
                                val = input(f"      Make invalid [YES/no] ? ")
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

        if newPlants>0:
            print(f"! Got {newPlants} new plants")

        print()
        return 0

    def createAccession(self):
        if self.pendingPlantImages:
            maxRow = self.hpsPlantsDB.workbook['Plants'].max_row
            accession = int(self.hpsPlantsDB.getValue('Plants', maxRow, 2)[1:]) # Number
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid==False or imageInfo.unknownProvenance==True:
                    continue
                maxRow    += 1
                accession += 1
                imageInfo.xlsxRow   = maxRow
                imageInfo.accession = accession

        if self.pendingGardenImages:
            maxRow = self.hpsGardensDB.workbook['Gardens'].max_row
            accession = int(self.hpsGardensDB.getValue('Gardens', maxRow, 2)[1:]) # Number
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid==False:
                    continue
                maxRow    += 1
                accession += 1
                imageInfo.xlsxRow   = maxRow
                imageInfo.accession = accession

        return 0

    def writeXlsxResults(self):
        print("Write xlsx results")
        print("------------------")
        now = datetime.datetime.now()

        if self.pendingPlantImages:
            print(f"* '{self.hpsPlantsDB.filename+self.hpsPlantsDB.extension}': master image database")
            print("  - Update", end="\r")
            validFiles = 0
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid==False or imageInfo.unknownProvenance==True:
                    continue
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 1, self.convertSpecialChar(imageInfo.getRHSName()))                # plantName
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 2, "P{:05d}".format(imageInfo.accession)) # HPSNo
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 3, imageInfo.getRHSNumber())              # RHSNo
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 4, imageInfo.getRHSStatus())              # RHSStatus
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 8, imageInfo.donor)                       # donor
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow, 9, imageInfo.dateAdded)                   # dateAdded
                self.hpsPlantsDB.setValue('Plants', imageInfo.xlsxRow,11, imageInfo.metaData)                    # extra information
                # Write data extracted from exif
                #imageInfo.extractExif()
                #print(imageInfo.exif)
                #if imageInfo.dateTimeOriginal:
                #    date = re.search(r'(\d{4}):(\d{2}):(\d{2})', imageInfo.dateTimeOriginal)
                #    if not date:
                #        date = re.search(r'(\d{4})-(\d{2})-(\d{2})', imageInfo.dateTimeOriginal)
                #    if date:
                #        if int(date.group(3)) != 0:
                #            self.hpsPlantsDB.setValue(imageInfo.xlsxRow, 'dateOriginal', date.group(3)+"/"+ date.group(2)+"/"+ date.group(1))
                #self.hpsPlantsDB.setValue(imageInfo.xlsxRow, 'fileSize', int(imageInfo.size)/1024)
                #self.hpsPlantsDB.setValue(imageInfo.xlsxRow, 'width', imageInfo.width)
                #self.hpsPlantsDB.setValue(imageInfo.xlsxRow, 'height', imageInfo.height)
                #if imageInfo.validateSize(False):
                #    self.hpsPlantsDB.setValue(imageInfo.xlsxRow, 'tooSmall', "X")
                #if imageInfo.make:
                #    self.hpsPlantsDB.setValue(imageInfo.xlsxRow, 'make', imageInfo.make)
                #if imageInfo.model:
                #    self.hpsPlantsDB.setValue(imageInfo.xlsxRow, 'model', imageInfo.model)
                # Update number of valid images
                #imageInfo.printPretty()
                validFiles += 1
            if validFiles == 0:
                print("  ! No data to write")
            else:
                newPath = self.uploadDir+'To_DropBox_'+now.strftime("%d%m%y")+'\\'+self.hpsPlantsDB.filename+self.hpsPlantsDB.extension
                print(f"  - Write to {newPath}")
                if not self.args.dryrun:
                    try:
                        self.hpsPlantsDB.save(newPath)
                    except PermissionError:
                        print(f"! Permission error writing to {newPath}")

            print(f"* '{self.imagelibDB.filename+self.imagelibDB.extension}': image database used by website")
            print("  - Update", end="\r")
            # Find where the P files (plants) change into X files (gardens)
            padd = 0
            for index in range(2, self.imagelibDB.workbook['active'].max_row):
                imageID = self.imagelibDB.getValue('active', index, 2) # Image ID
                if imageID.startswith("X"):
                    padd = index
                    break
            validFiles = 0
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid==False or imageInfo.unknownProvenance==True:
                    continue
                # Insert plants at end of P files
                self.imagelibDB.workbook['active'].insert_rows(padd)
                html = "<span RHS>" + imageInfo.rhsHtml[0] + "</span>"
                for index in range(1, len(imageInfo.rhsHtml)):
                    html += " && <span RHS>" + imageInfo.rhsHtml[index] + "</span>"
                #print(f"-  Insert image P{imageInfo.accession:05} in row {padd}")
                self.imagelibDB.setValue('active', padd, 1,  html) # Caption
                self.imagelibDB.setValue('active', padd, 2, f"P{imageInfo.accession:05}") # Image ID
                validFiles += 1
                padd += 1
            if validFiles == 0:
                print("! No data to write")
            else:
                newPath = self.uploadDir+'To_FTP_'+now.strftime("%d%m%y")+'\\'+self.imagelibDB.filename+self.imagelibDB.extension
                print(f"  - Write to {newPath}")
                if not self.args.dryrun:
                    try:
                        self.imagelibDB.save(newPath)
                    except PermissionError:
                        print("  ! Permission error writing to {}".format(self.imagelibDB.fileName))

            print(f"* '{self.generaDB.filename+self.generaDB.extension}': alphabetically sorted list of genera we have pictures of. Used by website.")
            print("  - Update", end="\r")
            existingGenus = self.generaDB.getColumn('active', 1)
            genusAdded = False
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid==False or imageInfo.unknownProvenance==True:
                    continue
                # Add the genus to the spreadsheet if it's not already there
                if len(imageInfo.rhsGenus) == 0:
                    continue;
                for index in range(len(imageInfo.rhsGenus)):
                    genus = imageInfo.rhsGenus[index].capitalize()
                    if genus in existingGenus:
                        #print(f"  - genus {genus} already in database")
                        continue
                    for rhsGenus in existingGenus:
                        # First 7 rows don't contain valid compare data
                        if existingGenus.index(rhsGenus)<8:
                            continue
                        if genus<rhsGenus:
                            print(f"  - Inserting new genus {genus}, family {imageInfo.rhsFamily[index].capitalize()} in row {existingGenus.index(rhsGenus)}")
                            self.generaDB.workbook['active'].insert_rows(existingGenus.index(rhsGenus)+1)
                            self.generaDB.setValue('active', existingGenus.index(rhsGenus)+1, 1, genus)
                            self.generaDB.setValue('active', existingGenus.index(rhsGenus)+1, 2, imageInfo.rhsFamily[index].capitalize())
                            genusAdded = True
                            # Recreate the list of existing genus as a new one has just been added
                            existingGenus = self.generaDB.getColumn('active', 1)
                            break
            if genusAdded:
                newPath = self.uploadDir+'To_FTP_'+now.strftime("%d%m%y")+'\\'+self.generaDB.filename+self.generaDB.extension
                print(f"  - Write to {newPath}")
                if not self.args.dryrun:
                    try:
                        self.generaDB.save(newPath)
                    except PermissionError:
                        print("  ! Permission error writing to {}".format(self.generaDB.fileName))
            else:
                print("  - No new genus added")

        if self.pendingGardenImages:
            print(f"* '{self.hpsGardensDB.filename+self.hpsGardensDB.extension}': master image database")
            print("  - Update", end="\r")
            validFiles = 0
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid==False:
                    continue
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 1, imageInfo.gardenName)                  # plantName
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 2, "X{:05d}".format(imageInfo.accession)) # HPSNo
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 3, imageInfo.donor)                       # donor
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 4, imageInfo.dateAdded)                   # dateAdded
                self.hpsGardensDB.setValue('Gardens', imageInfo.xlsxRow, 6, imageInfo.metaData)                    # extra information
                validFiles += 1
            if validFiles == 0:
                print("  ! No data to write")
            else:
                newPath = self.uploadDir+'To_DropBox_'+now.strftime("%d%m%y")+'\\'+self.hpsGardensDB.filename+self.hpsGardensDB.extension
                print(f"  - Write to {newPath}")
                if not self.args.dryrun:
                    try:
                        self.hpsGardensDB.save(newPath)
                    except PermissionError:
                        print(f"! Permission error writing to {newPath}")

            print(f"* '{self.imagelibDB.filename+self.imagelibDB.extension}': image database used by website")
            print("  - Update", end="\r")
            validFiles = 0
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid==False:
                    continue
                # Insert gardens at end of workbook
                row = self.imagelibDB.workbook['active'].max_row+1
                self.imagelibDB.setValue('active', row, 1, imageInfo.gardenName) # Caption
                self.imagelibDB.setValue('active', row, 2, f"X{imageInfo.accession:05}") # Image ID
                validFiles += 1
            if validFiles == 0:
                print("! No data to write")
            else:
                newPath = self.uploadDir+'To_FTP_'+now.strftime("%d%m%y")+'\\'+self.imagelibDB.filename+self.imagelibDB.extension
                print(f"  - Write to {newPath}")
                if not self.args.dryrun:
                    try:
                        self.imagelibDB.save(newPath)
                    except PermissionError:
                        print("  ! Permission error writing to {}".format(self.imagelibDB.fileName))

        print()
        return 0

    def copyToUpload(self):
        print("Copy images")
        print("-----------")

        if self.pendingPlantImages:
            # Copy plant images into upload directory for dropbox and remove GPS data
            print(f"* Copy plant  images to upload directory for dropbox  '{self.uploadDropBoxPlantsDir}' and remove GPS data")
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid==False or imageInfo.unknownProvenance==True:
                    continue
                startletter = imageInfo.getRHSName()[0]
                # Copy from pending to dropbox upload directory
                if not self.args.dryrun:
                    os.makedirs(self.uploadDropBoxPlantsDir+startletter, exist_ok=True)
                newFilename = self.uploadDropBoxPlantsDir+startletter+"\\"+self.convertSpecialChar(imageInfo.getRHSName())+" P{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                if not self.args.dryrun:
                    #print(f"  - copyfile({imageInfo.path}, {newFilename})")
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
            # Copy garden images into upload directory for dropbox and remove GPS data
            print(f"* Copy garden images to upload directory for dropbox  '{self.uploadDropBoxGardensDir}' and remove GPS data")
            for imageInfo in self.pendingGardensImageInfo:
                if imageInfo.valid==False:
                    continue
                # Copy from pending to dropbox upload directory
                if not self.args.dryrun:
                    os.makedirs(self.uploadDropBoxGardensDir, exist_ok=True)
                newFilename = self.uploadDropBoxGardensDir+"\\"+imageInfo.gardenName+" X{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                if not self.args.dryrun:
                    #print(f"  - copyfile({imageInfo.path}, {newFilename})")
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

        print(f"* Create thumbnails in {self.uploadFtpThumbsDir} (resize, auto orientate, remove exif, add watermark)")
        # The watermark will appear in the middle bottom, white, offset by 12 pixels.
        watermarkText = "gravity south fill white text 0,12 'Hardy Plant Society\\nwww.hardy-plant.org.uk'"

        if not self.args.dryrun:
            os.makedirs(self.uploadFtpThumbsDir, exist_ok=True)
        if self.pendingPlantImages:
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.valid==False or imageInfo.unknownProvenance==True:
                    continue
                startletter = imageInfo.getRHSName()[0]
                oldFilename = self.uploadDropBoxPlantsDir+startletter+"\\"+self.convertSpecialChar(imageInfo.getRHSName())+" P{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                oldFilename = oldFilename.replace(u'/', u'_')
                newFilename = self.uploadFtpThumbsDir+"P{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()

                # Convert image
                #print("  - Convert '{}'".format(oldFilename))
                #print("         to '{}'".format(newFilename))
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
                    if (out.returncode!=0):
                        print("  ! Error")
                        imageInfo.valid = False
                        continue;

            # Copy plant of unknown provenance into separate directory
            print(f"* Copy to upload directory for unknown provenance '{self.uploadUnknownProvenancePlantsDir}'")
            for imageInfo in self.pendingPlantsImageInfo:
                if imageInfo.unknownProvenance==False:
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
                if imageInfo.valid==False:
                    continue
                oldFilename = self.uploadDropBoxGardensDir+imageInfo.gardenName+" X{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()
                oldFilename = oldFilename.replace(u'/', u'_')
                newFilename = self.uploadFtpThumbsDir+"X{:05d}".format(imageInfo.accession)+imageInfo.getReformattedExtension()

                # Convert image
                #print("  - Convert '{}'".format(oldFilename))
                #print("         to '{}'".format(newFilename))
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
                    if (out.returncode!=0):
                        print("  ! Error")
                        imageInfo.valid = False
                        continue;

        print()
        return 0

    def printFinalResults(self):
        print("Final results")
        print("-------------")

        first = True
        for imageInfo in self.pendingPlantsImageInfo:
            if imageInfo.valid and imageInfo.unknownProvenance==False:
                if first:
                    print("* Images converted:")
                    first = False
                print("  - "+imageInfo.filename+imageInfo.extension)

        first = True
        for imageInfo in self.pendingPlantsImageInfo:
            if imageInfo.valid==False or imageInfo.unknownProvenance==True:
                if first:
                    print("* Images not converted:")
                    first = False
                print("  - "+imageInfo.filename+imageInfo.extension)

        print()
        return

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
        '--download',
        action='store_true',
        help="Download spreadsheets (Default: don't download)"
    )
    parser.add_argument(
        '--dryrun',
        action='store_true',
        help='Run without saving/creating any files'
    )
    parser.add_argument(
        '--fullAnalysis',
        action='store_true',
        help='Do a full analysis of the databases'
    )
    parser.add_argument(
        '--stats',
        help='Print out stats of database since given HPS number (e.g. P00001)'
    )
    args = parser.parse_args()

    # Construct the base class
    hps = CHPS(args)

    print()

    # Print out stats if requested
    if args.stats:
        if not re.search(r'P\d{5}', args.stats):
            print("Invalid start number");
            return 1
        hps.stats(args.stats)
        return 0

    # Do a full analysis of the databases
    if args.fullAnalysis:
        hps.fullAnalysis()
        return 0

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
    if hps.copyToUpload():
        return 1

    if hps.writeXlsxResults():
        return 1

    hps.printFinalResults()

    print("And finally")
    print("-----------")
    print("Copy to local disc:")
    print(f"  * Copy thumbnails from {hps.uploadFtpThumbsDir}")
    print(f"                    to   {hps.thumbsDir}")
    if hps.pendingPlantImages:
        print(f"  * Copy genera.csv from {hps.uploadFtpDir}")
        print(f"                    to   {hps.generaDB.path}")
    print(f"  * Copy imagelib.csv from {hps.uploadFtpDir}")
    print(f"                      to   {hps.imagelibDB.path}")
    if hps.pendingPlantImages:
        print(f"  * Copy new plant images from {hps.uploadDropBoxPlantsDir}")
        print(f"                          to   {hps.plantsDir}")
        print(f"  * Copy 'HPS Images - Plants.xlsx' from {hps.uploadDropBoxPlantsDir}")
        print(f"                                    to   {hps.hpsPlantsDB.path}")
    if hps.pendingGardenImages:
        print(f"  * Copy new garden images from {hps.uploadDropBoxGardensDir}")
        print(f"                           to   {hps.gardensDir}")
        print(f"  * Copy 'HPS Images - Gardens.xlsx' from {hps.uploadDropBoxPlantsDir}")
        print(f"                                    to   {hps.hpsPlantsDB.path}")
    print()
    print("Copy to ftp:")
    print(f"  * Copy thumbnails from {hps.uploadFtpThumbsDir}")
    print(f"                    to   ftp://images@hardy-plant.org.uk:/catalog/library/thumbs")
    if hps.pendingPlantImages:
        print(f"  * Copy genera.csv from {hps.uploadFtpDir}")
        print(f"                    to   ftp://docsftp@hardy-plants.org.uk:/plants if available")
    print(f"  * Copy imagelib.csv from {hps.uploadFtpDir}")
    print(f"                      to   ftp://docsftp@hardy-plants.org.uk:/plants")
    print()
    print("Copy to dropbox:")
    if hps.pendingPlantImages:
        print(f"  * Copy new plant images from {hps.uploadDropBoxPlantsDir}")
        print("                          to   https://www.dropbox.com/home/Family%20Room/Image%20Library/Images%20-%20plants")
        print(f"  * Copy 'HPS Images - Plants.xlsx' from {hps.uploadDropBoxPlantsDir}")
        print("                                    to   https://www.dropbox.com/home/Family%20Room/Image%20Library")
    if hps.pendingGardenImages:
        print(f"  * Copy new garden images from {hps.uploadDropBoxGardensDir}")
        print("                           to   https://www.dropbox.com/home/Family%20Room/Image%20Library/Images%20-%20Gardens")
        print(f"  * Copy 'HPS Images - Gardens.xlsx' from {hps.uploadDropBoxPlantsDir}")
        print("                                     to   https://www.dropbox.com/home/Family%20Room/Image%20Library")

    return 0

if __name__ == "__main__":
    ret = main()
    sys.exit(ret)

sys.exit(1)