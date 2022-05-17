# Import everything needed to edit video clips
from asyncio import threads
from concurrent.futures import thread
from moviepy.editor import *
import shutil
import openpyxl
import os

# Getting input data of clip name and timestamp locations
#####
# Location of spreadsheet
loc = ('C:/Users/keelo/source/repos/pgmVideoClips/footage/014_A_annotations.xlsx')
# Open workbook
wb_obj = openpyxl.load_workbook(loc, data_only=True)
sheet_obj = wb_obj.active
# Starting row is 3 after header, columns stay the same
rowCount = 3

## Retreiving Clip ##
# loading video gfg
vidName = input("Enter video name and extension: ")


for i in range(rowCount, sheet_obj.max_row+1):
    rowCount = i
    
    # Set start and end times of timestamp
    tStart = sheet_obj.cell(row=rowCount, column=4)
    tEnd = sheet_obj.cell(row=rowCount, column=5)
    
    # Set lcoation of clip flags
    cfloc = sheet_obj.cell(row=rowCount, column=1)
    
    if (sheet_obj.cell(row=rowCount, column=1).value == "N"):

        clip = VideoFileClip(
            'C:/Users/keelo/source/repos/pgmVideoClips/footage/' + vidName)

        # getting only video between timestamp 1 and 2 as strings
        clip = clip.subclip(str(tStart.value), str(tEnd.value))
        # Save clip
        clipName = (str(sheet_obj.cell(row=rowCount, column=2).value) +
                    "_" + (str(tStart.value))[:-4].replace(":", ".") + ".mp4")
        clip.write_videofile(clipName)

        #####################
        # Move the clip to correct folder
        destination = 'C:/Users/keelo/source/repos/pgmVideoClips/Clips'
        dest = shutil.move(
            'C:/Users/keelo/source/repos/pgmVideoClips/' + clipName, destination, copy_function=shutil.copytree)

        # Change Clip Flag
        cfloc.value = "Y"
        wb_obj.save('C:/Users/keelo/source/repos/pgmVideoClips/footage/014_A_annotations.xlsx')
    else:
         continue
    
print("End of spreedsheet!")

### Add to filename of spreadsheet feature ###
