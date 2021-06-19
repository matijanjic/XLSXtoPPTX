
#  A script that takes an existing excel file, reads the required columns that hold the info
# (word, sentence, word audio, sentence audio, sentence picture etc.) and arranges it all into a
# pptx presentation. The use case is pretty specific, my wife needed it for her phd experiment.
# Audio files were generated using another script I made that uses the Google text-to-speech
# (gTTS) library that read the words and sentences from a csv file and exported them as .mp3's

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from collections import defaultdict
from slideshow import *

# returns a dictionary of lists where each letter key has a value that is a list of values from that column
# in the worksheet
def getDictFromXlsx(xlsxFile, rowStart, rowEnd, colStart, colEnd, **kwargs):

    wb = load_workbook(xlsxFile, data_only=True)
    ws = wb.active

    # goes through the workbook and returns a dictionary
    wsDict = defaultdict(list)
    for row in range(rowStart, rowEnd):
        for column in range(colStart, colEnd):
            colLetter = get_column_letter(column)
            if colLetter in kwargs.values():
                wsDict[colLetter].append(ws.cell(row=row, column=column).value)
    return wsDict


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

    # starting and ending points in the workbook
    rowStart = 8
    rowEnd = 25
    colStart = 6
    colEnd = 11

    # create a new instance of the SlideShow class that takes the 
    # layout file (in this case a 16x9 ratio file) and the slide layout (further explained in the 
    # python-pptx documentation https://python-pptx.readthedocs.io/en/latest/user/slides.html)
    slideShow = SlideShow('16x9.pptx', 6)

    # use the getDictFromXlsx function that returns a filled out dictionary where the keys are the column letters
    # and the values are lists that contain the data in those columns
    xlsxDict = getDictFromXlsx(xlsxFile, rowStart, rowEnd + 1, colStart, colEnd,
                               wordTextCol='F', wordSoundCol='G', sentenceSoundCol='I', sentencePictureCol='J')

    # -- MAIN SLIDE LAYOUT --#

    # for each row create a following slide layout:
    for i in range(rowEnd - rowStart + 1):
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
