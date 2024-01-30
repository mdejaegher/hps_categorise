# Prepare your images for the HPS Image Library

## Requirements and installation
The script relies on a couple of programs that have to be installed to work properly:
1. *Python*: this is needed to run the scripts and is the language the scripts are written in. For Windows, I found the easiest way to install Python was by going to the command promt and type 'Python'. This will open the Microsoft Store from where you can download the latest version. If you're familiar with Linux then I recommend using WSL2 (https://ubuntu.com/tutorials/install-ubuntu-on-wsl2-on-windows-10) which will install Python for you.
2. *ImageMagick*: this is used to create the thumbnails by reducing the size and adding a watermark. Can be downloaded from https://imagemagick.org/script/download.php. This will download an installer which you will need to run.
3. *exiftool*: this is used for security by removing all the data hidden inside photos, e.g. lat/long. Can be downloaded from https://exiftool.org/. You will need to install this yourself so is a bit more involved.
   * You should have downloaded a file that looks like 'exiftool-12.70.zip' (probably in the Downloads directory)
   * You can extract this by rightclicking and use 'Extract All...' which will create a new folder 'exiftool-12.70' which should contain just one file 'exiftool(-k).exe'.
   * You now need to go to 'C:\Program Files' and create a new folder, e.g. 'exiftool' and move the file into that folder.
   * Once it's there, rename it to 'exiftool.exe'.
   * Open Explorer, right click on 'This PC' and open 'Properties...'
   * Go to 'Advanced System Settings' which should open a 'System Properties' window
   * Click on 'Environment Variables' which should open another window.
   * In 'System variabes', you should see 'Path'. Click on that and then click on 'Edit...'
   * Click on 'New' and type 'C:\Program Files\exiftool'
   * Click on 'OK' for all the windows to close them again.
4. *WinMerge*: if you are maintaining a backup of all the images then I would recommend using WinMerge (https://winmerge.org/?lang=en). It's an easy tool to compare two directories (or directory structures) and it will tell you which images are different or missing between the two.

The script also needs a number of spreadsheets to allow cross referencing with the RHS database and to add to our database. This is available via https://github.com/mdejaegher/hps_categorise :
* *genera.csv*: this is the database used by the website to list all the genus. I've never had to update this spreadsheet.
* *HPS Images - Gardens.xlsx*: list of all HPS images of gardens with lots of details
* *HPS Images - Plants.xlsx*: list of all HPS images of plants with lots of details
* *imagelib.csv*: this is the database used by the website to list all the images and their plant names. It contains the properly formatted name of the plant and the thumbnail number. There are a lot of special characters in this file so be careful which editor you use if you want to make some changes as not all editors deal well with the characters.
* RHS_dataset.xlsx: this is the RHS database of known plants. This database is under license and therefore not to be shared with anyone else!
* *RHS Names - Hardy Plant Society - Data Sharing Sep 2023.xlsx/RHS Names - HPS - Sep 2023 - Reduced - Unlocked.xlsx*: the former is the latest version of the RHS database. This is locked by a password. In order to make it easier/faster for the tools to read the spreadsheet, I created the latter which only contains the data we need and isn't locked. Again, both are under license and therefore not to be shared with anyone else!


## How to use the script
The main script is called `prepareImages.py` and relies on two additional files `CImageInfo.py` and `CSpreadSheet.py` to function properly.

### First time running the script
There are a few things the script needs to know before it can start:
* *Where do the images live?* All the images (both classified and pending) live in the same base directory. The base directory is where all images can be found from and is defined in `self.baseDir` in `prepareImages.py`. It's the directory where you would expect to find the subdirectories `Gardens` and `Plants`.
From there, it can then find the directories containing all the plant images, garden images, pending plant images, pending garden images and the upload directory.
* *Where do the scripts live?* In order to find all the scripts and spreadsheet, you need to set `self.gitHubDir` in `prepareImages.py` to the directory where you can find the scripts/spreadsheets.

You should now be able to run the script for the first time from the command line like this:
```
python prepareImages.py
```
It should at least finish with the validation of tools and input. If something is missing then it should let you know what's missing and try to suggest what to do.

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
