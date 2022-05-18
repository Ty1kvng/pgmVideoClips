# Import everything needed to edit video clips
from importlib.resources import path
import inspect
from modulefinder import packagePathMap
from moviepy.editor import VideoFileClip
import shutil
import openpyxl
import os

# Getting input data of clip name and timestamp locations
#####
# Location of spreadsheet
loc = input("Enter absolute path of spreadsheet: ")
# Open workbook
wb_obj = openpyxl.load_workbook(loc, data_only=True)
sheet_obj = wb_obj.active
# Starting row is 3 after header, columns stay the same
rowCount = 3

## Retreiving Clip ##
# loading video gfg
vidName = input("Enter video name and extension:  ")


for i in range(rowCount, sheet_obj.max_row+1):
    rowCount = i

    # Set start and end times of timestamp
    tStart = sheet_obj.cell(row=rowCount, column=4)
    tEnd = sheet_obj.cell(row=rowCount, column=5)

    # Set location of clip flags
    cfloc = sheet_obj.cell(row=rowCount, column=1)

    if (sheet_obj.cell(row=rowCount, column=1).value == "N"):

        clip = VideoFileClip(
            os.path.abspath('Main/') + "\\" + vidName)

        # getting only video between timestamp 1 and 2 as strings
        clip = clip.subclip(str(tStart.value), str(tEnd.value))
        # Save clip
        clipName = (str(sheet_obj.cell(row=rowCount, column=2).value) +
                    "_" + (str(tStart.value))[:-4].replace(":", ".") + ".mp4")
        clip.write_videofile(clipName)
        
        # Save new clipName to spreadsheet clipName
        fileNameLoc = sheet_obj.cell(row=rowCount, column=2)
        fileNameLoc.value = clipName


        #####################
        # Move the clip to correct folder
        pgmsource = (os.path.dirname(os.path.abspath(
            inspect.getfile(inspect.currentframe())))) + "\\"

        destination = os.path.abspath('Clips/')
        dest = shutil.move(pgmsource + clipName, destination,
                           copy_function=shutil.copytree)

        # Change Clip Flag
        cfloc.value = "Y"
        wb_obj.save(loc)
    else:
        continue

print("End of spreedsheet!")
