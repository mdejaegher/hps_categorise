#!/usr/bin/python
import json
import os
import subprocess


class CImageInfo:
    def __init__(self, path, verbose):
        # Copy arguments
        self.path = path
        self.verbose = verbose

        # Derived variables
        self.dirname = os.path.dirname(path)
        self.filename = os.path.basename(os.path.splitext(path)[0])
        self.extension = os.path.splitext(path)[1]

        # To be filled in later
        self.valid = True
        self.unknownProvenance = False
        self.width = None
        self.height = None

    def __str__(self):
        return (f"<CImageInfo path: {self.path}, filename: {self.filename}, " +
                f"extension: {self.extension}, size: {self.size}, md5: " +
                f"{self.md5}, format: {self.format}, width: {self.width}, " +
                f"height: {self.height}, fstop: {self.fstop}, exposure: " +
                f"{self.exposure}, ISO: {self.ISO}, make: {self.make}, model: " +
                f"{self.model}, dateTimeOriginal: {self.dateTimeOriginal}>")

    def validateSize(self):
        # Had cases where two images with similar dimensions where one was 3
        # times larger than the other. Zooming in did show one had better detail
        # than the other. Adding check for (arbitrary) size: if it's larger than
        # 1.5MB the it's OK
        if os.path.getsize(self.path) > 1.5*1024*1024:
            return 0

        # This is based on the assumption that we want to print out 5"x7" at
        # near photo quality which needs an image size of 1536x1180
        # (see http://www.urban75.org/photos/print.html)
        if self.width >= 1536 and self.height >= 1180:
            return 0
        if self.width >= 1180 and self.height >= 1536:
            return 0
        return 1

    def extractExif(self):
        # Extract exif
        out = subprocess.Popen(["magick",
                                "convert",
                                self.path,
                                "json:"], stdout=subprocess.PIPE).communicate()[0]

        exif = json.loads(out.decode(errors='ignore'))
        self.width = exif[0]['image']['geometry']['width']
        self.height = exif[0]['image']['geometry']['height']

        return 0

    def getReformattedExtension(self):
        extension = self.extension.lower()
        if extension == 'jpeg':
            extension = 'jpg'
        return extension


# Information class for pending HPS images
class CPendingImageInfo(CImageInfo):
    def __init__(self, path):
        CImageInfo.__init__(self, path, False)
        self.donor = None
        self.dateAdded = None
        self.metaData = None
        self.email = None
        self.gardenName = None
        self.xlsxRow = 0
        self.accession = 0
        self.rhsNumbers = []
        self.rhsFamily = []
        self.rhsGenus = []
        self.rhsSpecies = []
        self.rhsCultivar = []
        self.rhsNames = []
        self.rhsHtml = []
        self.valid = True

        self.extractExif()

    def __str__(self):
        return f"<CPendingImageInfo valid:{self.valid}, path:{self.path}>"

    def printPretty(self):
        print(f"    - HPS Donor:    '{self.donor}'")
        print("      HPS accession number: 'P/X{:05d}'".format(self.accession))
        print(f"      HPS xlsx row: '{self.xlsxRow}'")
        print(f"      RHS number:   '{self.rhsNumbers}'")
        print(f"      RHS family:   '{self.rhsFamily}'")
        print(f"      RHS genus:    '{self.rhsGenus}'")
        print(f"      RHS species:  '{self.rhsSpecies}'")
        print(f"      RHS cultivar: '{self.rhsCultivar}'")
        print(f"      RHS name:     '{self.rhsNames}'")
        print(f"      RHS html:     '{self.rhsHtml}'")

    def getRHSName(self):
        rhsNameString = ""
        numRHSNames = len(self.rhsNames)
        if numRHSNames > 0:
            rhsNameString = self.rhsNames[0]
            if numRHSNames > 1:
                for x in range(1, len(self.rhsNames)):
                    if self.rhsNames[x] != "":
                        rhsNameString += " && " + self.rhsNames[x]
        return rhsNameString

    def getRHSNumber(self):
        rhsNumberString = None
        numRHSNumbers = len(self.rhsNumbers)
        if numRHSNumbers > 0:
            rhsNumberString = self.rhsNumbers[0]
            if numRHSNumbers > 1:
                rhsNumberString = str(rhsNumberString)
                for x in range(1, len(self.rhsNumbers)):
                    if self.rhsNumbers[x] != 0:
                        rhsNumberString += " && " + str(self.rhsNumbers[x])
        return rhsNumberString
