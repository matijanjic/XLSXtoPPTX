
#  A script that takes an existing excel file, reads the required columns that hold the info
# (word, sentence, word audio, sentence audio, sentence picture etc.) and arranges it all into a
# pptx presentation. The use case is pretty specific, my wife needed it for her phd experiment.
# Audio files were generated using another script I made that uses the Google text-to-speech
# (gTTS) library that read the words and sentences from a csv file and exported them as .mp3's

from openpyxl import load_workbook
from collections import defaultdict
from slideshow import *

# returns a dictionary of lists where each letter key has a value that is a list of values from that column
# in the worksheet. 
def getDictFromXlsx(xlsxFile, colList):

    wb = load_workbook(xlsxFile, data_only=True)
    ws = wb.active

    values = defaultdict(list)

    for letter in colList:
        column = ws[letter]
        for cell in column:
            if cell.value != None:
                print(cell.column_letter)
                values[cell.column_letter].append(cell.value)
    
    return values
        
def main():
    # some constants declared here
    xlsxFile = 'excel\spreadsheet_gorilla_learning_pictures_-all.xlsx'
    wordSoundFolder = 'sounds\words\\'
    sentenceSoundFolder = 'sounds\sentences\\'
    pictureFolder = 'images\pictures_learning\\'

    # these could be changed to what ever fits your needs, but are column letters that hold the data for the powerpoint
    wordTextCol = 'F'
    wordSoundCol = 'G'
    sentenceSoundCol = 'I'
    sentencePictureCol = 'J'
    # list of all the columns so they can be searched more easily. If the number of columns were bigger, it would pay off to automate it
    colList = ['F', 'G', 'I', 'J']
    
    # create a new instance of the SlideShow class that takes the width and the height in inches
    # and the and the slide layout (further explained in the python-pptx documentation 
    # https://python-pptx.readthedocs.io/en/latest/user/slides.html)
    # e.g. SlideShow(16, 9, 6) creates a 16x9 presentation file with a blank slide template (layout number 6)
    slideShow = SlideShow(16, 9, 6)

    # use the getDictFromXlsx function that returns a filled out dictionary where the keys are the column letters
    # and the values are lists that contain the data in those columns. 
    xlsxDict = getDictFromXlsx(xlsxFile, colList)
    numberOfRows = len(list(xlsxDict.values())[0])
    
    # -- MAIN SLIDE LAYOUT --#

    # for each row create a following slide layout:
    for i in range(numberOfRows):
        
        slideShow.addSlide()
        slideShow.addText(80, 'X', 4, 1)
        slideShow.addSlide()
        slideShow.addText(44, xlsxDict[wordTextCol][i], 4, 1)
        slideShow.addSound(wordSoundFolder + xlsxDict[wordSoundCol][i])
        slideShow.addSlide()
        slideShow.addPicture(pictureFolder + xlsxDict[sentencePictureCol][i], 400)
        slideShow.addSound(sentenceSoundFolder + xlsxDict[sentenceSoundCol][i])
    # and save it
    slideShow.save('test.pptx')


if __name__ == '__main__':
    main()
