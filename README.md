# hps_categorise
Prepare your images for the HPS Image Library

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
