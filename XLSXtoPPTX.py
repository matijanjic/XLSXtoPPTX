
#  A script that takes an existing excel file, reads the required columns that hold the info 
# (word, sentence, word audio, sentence audio, sentence picture etc.) and arranges it all into a
# pptx presentation. The use case is pretty specific, my wife needed it for her phd experiment.
# Audio files were generated using another script I made that uses the Google text-to-speech 
# (gTTS) library that read the words and sentences from a csv file and exported them as .mp3's 

from pptx import Presentation
from pptx.util import Inches, Pt
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image

def main():

    # import presentation layout as prs
    prs = Presentation('16x9.pptx')

    # slide layout and constants config
    slide_layout = prs.slide_layouts[6]
    width = height = Inches(1)
    picHeight = Inches(3)
    picLeft = Inches(5)

    # column letter assigment as in loaded excel table
    wordTextColumn = 'F'
    wordSoundColumn = 'G'
    sentenceSoundColumn = 'I'
    sentencePictureColumn = 'J'

    # load existing workbook file as data only and set it as active
    wb = load_workbook('excel\spreadsheet_gorilla_learning_pictures.xlsx', data_only=True)
    ws = wb.active

    # counter to iterate over all the rows in the workbook
    for row in range(8,281):

        # this part of the program creates a slide configuration based on what is needed on the slide
        slide = prs.slides.add_slide(slide_layout)
        txBox = slide.shapes.add_textbox(Inches(6), Inches(2.5), width, height)
        tf = txBox.text_frame

        # add large letter X in the middle for experiment fixation
        p = tf.add_paragraph()
        p.text = "X"
        p.font.size = Pt(80)
        p.font.name = 'Calibri light'
        
        # iterate over the columns
        for col in range(6,11):

            # an if-else checking the column letter using the get_column_letter() method. Case would be better suited.
            if get_column_letter(col) == wordTextColumn:

                # add a word slide - word is selected from the current workbook cell as wordText
                slide = prs.slides.add_slide(slide_layout)
                wordText = ws.cell(row=row, column=col).value
                txBox = slide.shapes.add_textbox(Inches(5.5), Inches(2.8), width, height)
                tf = txBox.text_frame

                p=tf.add_paragraph()
                p.text = wordText.capitalize()
                p.font.size = Pt(44)
                p.font.name = 'Calibri light'

            elif get_column_letter(col) == wordSoundColumn:

                # add sound to the slide above using the filenames given in the workbook at the word sound column.
                wordSound = ws.cell(row=row, column=col).value
                sound = slide.shapes.add_movie("sounds\words\\" + wordSound, Inches(0.5), Inches(0.5), 0, 0)

                # a bit of xml editing using the etree method from the lxml module making sound autoplay possible. 
                # Solution found at https://github.com/scanny/python-pptx/issues/427. Thanks to iota-pi for the solution!
                tree = sound._element.getparent().getparent().getnext().getnext()
                timing = [el for el in tree.iterdescendants() if etree.QName(el).localname == 'cond'][0]
                timing.set('delay', '0')

            elif get_column_letter(col) == sentenceSoundColumn:

                # new slide added and sound played same as above, only sentence audio is used this time
                slide = prs.slides.add_slide(slide_layout)
                sentenceSound = ws.cell(row=row, column=col).value
                
                sound = slide.shapes.add_movie("sounds\sentences\\" + sentenceSound, Inches(0.5), Inches(0.5), 0, 0)
                tree = sound._element.getparent().getparent().getnext().getnext()
                timing = [el for el in tree.iterdescendants() if etree.QName(el).localname == 'cond'][0]
                timing.set('delay', '0')

            elif get_column_letter(col) == sentencePictureColumn:

                # adding a picture on the slide above. Filename is again provided by the workbook
                sentencePicture = ws.cell(row=row, column=col).value

                # use of the PIL library, opening each picture to check the orientation using a simple if else statement and
                # adjusting the picture origin on the x axis (picLeft, 0 is completely left) and the picture height - horizontal images
                # are given more height then the portrait ones so it fills up  the screen more efficiently and are moved more to the left.
                img = Image.open("images\pictures_learning\\" + sentencePicture)
                width, height = img.size
                
                if height >= width:
                    picHeight = Inches(3)
                    picLeft = Inches(5)
                else:
                    picHeight = Inches(6)
                    picLeft = Inches(3.2)
                slide.shapes.add_picture("images\pictures_learning\\" + sentencePicture, picLeft, Inches(1.5), picHeight)

    # presentation is saved.
    prs.save("learning pictures.pptx")

if __name__ == '__main__':
    main()