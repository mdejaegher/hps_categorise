#!/usr/bin/python
from CSpreadSheet import CSpreadSheet

import argparse
import os
import re
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
        print("Analyse directories")
        print("-------------------")
        print("- Check file in correct directory")
        allCorrect = True
        for plantsLetterDir in os.listdir(self.plantsDir):
            for filename in os.listdir(self.plantsDir+plantsLetterDir):
                if filename.startswith("x "):
                    if filename[2] != plantsLetterDir:
                        allCorrect = False
                        print(f"  ! File '{filename}' is in directory '{plantsLetterDir}' but should be in '{filename[2]}'")
                else:
                    if not filename.startswith(plantsLetterDir):
                        allCorrect = False
                        print(f"  ! File '{filename}' is in directory '{plantsLetterDir}' but should be in '{filename[0]}'")
        if allCorrect:
            print("    All plants are in correct directory")
        print()

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

################################################################################

def main():
    # Process the arguments
    parser = argparse.ArgumentParser(
        description='Stats on images.',

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

if __name__ == "__main__":
    ret = main()
    sys.exit(ret)

sys.exit(1)