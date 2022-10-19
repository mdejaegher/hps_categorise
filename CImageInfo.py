#!/usr/bin/python
import hashlib
import json
import os
import re
import shutil
import subprocess

class CImageInfo:
    def __init__(self, path, verbose):
        # Copy arguments
        self.path    = path
        self.verbose = verbose
        
        # Derived variables
        self.dirname   = os.path.dirname(path)
        self.filename  = os.path.basename(os.path.splitext(path)[0])
        self.extension = os.path.splitext(path)[1]
        
        # To be filled in later
        self.valid             = True
        self.unknownProvenance = False
        #self.size             = None
        #self.md5              = None
        #self.exif             = None
        #self.format           = None
        self.width             = None
        self.height            = None
        #self.fstop            = None
        #self.exposure         = None
        #self.ISO              = None
        #self.make             = None
        #self.model            = None
        #self.dateTimeOriginal = None
        #self.hasGPS           = False

        #self.extractExif()

    def __str__(self):
        return (f"<CImageInfo path: {self.path}, filename: {self.filename}, " +
               f"extension: {self.extension}, size: {self.size}, md5: " +
               f"{self.md5}, format: {self.format}, width: {self.width}, " +
               f"height: {self.height}, fstop: {self.fstop}, exposure: " +
               f"{self.exposure}, ISO: {self.ISO}, make: {self.make}, model: " +
               f"{self.model}, dateTimeOriginal: {self.dateTimeOriginal}>")

    #def getSize(self):
    #    if not self.size:
    #        self.size = os.path.getsize(self.path)
    #    return self.size

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
        if self.width>=1536 and self.height>=1180:
            return 0
        if self.width>=1180 and self.height>=1536:
            return 0
        return 1

    #def getFileHash(self):
    #    if not self.md5:
    #        hasher = hashlib.md5()
    #        with open(self.path, 'rb') as afile:
    #            buf = afile.read()
    #            hasher.update(buf)
    #        self.md5 = hasher.hexdigest()
    #    return self.md5

    def extractExif(self):
        # Extract exif
        out = subprocess.Popen(["magick",
                                "convert",
                                self.path,
                                "json:"], stdout=subprocess.PIPE).communicate()[0]

        exif = json.loads(out.decode(errors='ignore'))
        #print(self.exif)
        #self.format = self.exif[0]['image']['format']
        self.width  = exif[0]['image']['geometry']['width']
        self.height = exif[0]['image']['geometry']['height']

        #exifProperties = self.exif[0]['image']['properties']
        #if 'exif:FNumber' in exifProperties:
        #    fstop = re.search(r'(\d+)/(\d+)', exifProperties['exif:FNumber'])
        #    if fstop:
        #        fstop1 = float(fstop.group(1))
        #        fstop2 = float(fstop.group(2))
        #        self.fstop = f"f/{fstop1/fstop2}"
        #    else:
        #        if self.verbose: print(f"    Don't know what to do with fStop {exifProperties['exif:FNumber']} in {self.filename}")
        #if 'exif:ExposureTime' in exifProperties:
        #    exposure = re.search(r'(\d+)/(\d+)', exifProperties['exif:ExposureTime'])
        #    if exposure:
        #        exposure1 = int(exposure.group(1))
        #        exposure2 = int(exposure.group(2))
        #        time = int(exposure2/exposure1)
        #        self.exposure = f"1/{time}"
        #if 'exif:PhotographicSensitivity' in exifProperties:
        #    self.ISO = exifProperties['exif:PhotographicSensitivity']
        #if 'exif:Make' in exifProperties:
        #    self.make = exifProperties['exif:Make']
        #if 'exif:Model' in exifProperties:
        #    self.model = exifProperties['exif:Model']
        #if 'exif:DateTimeOriginal' in exifProperties:
        #    date = re.search(r'(\d{4})[:-](\d{2})[:-](\d{2})', exifProperties['exif:DateTimeOriginal'])
        #    if date:
        #        if int(date.group(3)) != 0:
        #            self.dateTimeOriginal = date.group(3)+"/"+ date.group(2)+"/"+ date.group(1)
        #if not self.dateTimeOriginal and 'date:create' in exifProperties:
        #    date = re.search(r'(\d{4})[:-](\d{2})[:-](\d{2})', exifProperties['date:create'])
        #    if date:
        #        if int(date.group(3)) != 0:
        #            self.dateTimeOriginal = date.group(3)+"/"+ date.group(2)+"/"+ date.group(1)
        #if 'exif:GPSAltitude' in exifProperties:
        #    self.hasGPS = True

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
        self.donor       = None
        self.dateAdded   = None
        self.metaData    = None
        self.email       = None
        self.gardenName  = None
        self.xlsxRow     = 0
        self.accession   = 0
        self.rhsNumbers  = []
        self.rhsFamily   = []
        self.rhsGenus    = []
        self.rhsSpecies  = []
        self.rhsCultivar = []
        self.rhsStatus   = []
        self.rhsNames    = []
        self.rhsHtml     = []
        self.valid       = True

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
        print(f"      RHS status:   '{self.rhsStatus}'")
        print(f"      RHS name:     '{self.rhsNames}'")
        print(f"      RHS html:     '{self.rhsHtml}'")

    def getRHSName(self):
        rhsNameString = ""
        numRHSNames = len(self.rhsNames)
        if numRHSNames > 0:
            rhsNameString = self.rhsNames[0]
            if numRHSNames>1:
                for x in range(1, len(self.rhsNames)):
                    if self.rhsNames[x] != "":
                        rhsNameString += " && " + self.rhsNames[x]
        return rhsNameString

    def getRHSNumber(self):
        rhsNumberString = None
        numRHSNumbers = len(self.rhsNumbers)
        if numRHSNumbers > 0:
            rhsNumberString = self.rhsNumbers[0]
            if numRHSNumbers>1:
                rhsNumberString = str(rhsNumberString)
                for x in range(1, len(self.rhsNumbers)):
                    if self.rhsNumbers[x] != 0:
                        rhsNumberString += " && " + str(self.rhsNumbers[x])
        return rhsNumberString

    def getRHSStatus(self):
        rhsStatusString = ""
        numRHSStatus = len(self.rhsStatus)
        if numRHSStatus > 0:
            rhsStatusString = self.rhsStatus[0]
            if numRHSStatus>1:
                for x in range(1, len(self.rhsStatus)):
                    if self.rhsStatus[x] != 0:
                        rhsStatusString += " && " + self.rhsStatus[x]
        return rhsStatusString
