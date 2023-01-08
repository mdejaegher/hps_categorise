# Prepare your images for the HPS Image Library

## Requirements
The script relies on a couple of programs that have to be installed to work properly:
1. Python: this is needed to run the scripts and is the language the scripts are written in. Can be downloaded from https://www.python.org/downloads/windows/
2. ImageMagick: this is used to create the thumbnails by reducing the size and adding a watermark. Can be downloaded from https://imagemagick.org/script/download.php
3. exiftool: this is used for security by removing all the data hidden inside photos, e.g. lat/long. Can be downloaded from https://exiftool.org/

If you are maintaining a backup of all the images then I would recommend using WinMerge (https://winmerge.org/?lang=en). It's an easy tool to compare two directories (or directory structures) and it will tell you which images are different or missing between the two.

It also needs a number of spreadsheets to allow cross referencing with the RHS database and to add to our database:
* genera.csv: this is the database used by the website to list all the genus. I've never had to update this spreadsheet.
* HPS Images - Gardens.xlsx: list of all HPS images of gardens with lots of details
* HPS Images - Plants.xlsx: list of all HPS images of plants with lots of details
* imagelib.csv: this is the database used by the website to list all the images and their plant names. It contains the properly formatted name of the plant and the thumbnail number. There are a lot of special characters in this file so be careful which editor you use if you want to make some changes as not all editors deal well with the characters.
* RHS_dataset.xlsx: this is the RHS database of known plants. This database is under license and therefore not to be shared with anyone else!

In order to uploaded the thumbnails and spreadsheets to our web page, I recommend you use WinSCP. Can be downloaded from https://winscp.net/eng/download.php

## How to use the script
The main script is called `prepareImages.py` and relies on two additional scripts `CImageInfo.py` and `CSpreadSheet.py` to function properly.

### First time running the script
There are a few things the script needs to know before it can start:
* Where do the images live? All the images (both classified and pending) live in the same base directory.

The base directory where all images can be found from is defined in self.baseDir

It can then find the directories containing all the plant images, garden images,
pending plant images, pending garden images and the upload directory from that
base directory.

### Preparing the images
Image file names are in the form of

    <plant name> <RHS number> <donor> <year> <misc>.jpg

If multiple plants, the form is

    <plant name> <RHS number> <donor> <year> && <plant name> <RHS number> <donor> <year> <misc>.jpg

If 'RHS number' not known, use '0'
'donor' can be 'unknown' or 'anonymous'
If 'year' not specified or 0 then the current date will be used instead.

If you have two different pictures of the same plant of the same donor then I
would suggest adding an extra space before/after the RHS number

### Running the script
To run the script, start the command prompt and run

    python prepareImages.py

The script will go through a number of steps to classify images:
* The first step is checking if the tools the script needs have been downloaded
  ('magick' and 'exiftool'). If it can't find it then it will tell where to 
  download it from.
* The second step is validating the input: are all the directories it expects
  there, are all the spreadsheets it expects there and are they in the format it
  needs.
* It will import all the existing and pending images in the next step

### Archiving the results
