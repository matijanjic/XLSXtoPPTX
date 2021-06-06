from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image
import PIL
import io

class SlideShow:
    
    def __init__(self, layoutFile, layoutNumber):
        self.layoutFile = layoutFile
        self.layoutNumber = layoutNumber
        self.ss = Presentation(layoutFile)
        self.slideLayout = self.ss.slide_layouts[layoutNumber]

    # adds a slide
    def addSlide(self):
        self.slide = self.ss.slides.add_slide(self.slideLayout)

    # adds a text object on the slide
    def addText(self, fontSize, text, width, height, left = 'center', top='center',):
        if left == 'center' and top == 'center':
            left = self.ss.slide_width / 2
            top = self.ss.slide_height / 2
            left = left - width / 2
            top = top - height / 2
        txBox = self.slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        txBox.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        p=tf.add_paragraph()
        p.text = text
        p.font.size = Pt(fontSize)

    # add a picture on the slide
    def addPicture(self, imgFile, maxSize, left = 'center', top = 'center'):

        # open the image
        img = Image.open(imgFile)

        # depending on the picture orientation, set the longest side to the maxSize argument (in pixels)
        width, height = img.size
        ratio = height / width
        if height > width:
            height = maxSize
            width = int(ratio * maxSize)
        else:
            width = maxSize
            height = int(ratio * maxSize)

        # resize the image using the calculations above and save it to img_resized variable
        img_resized = img.resize([width, height], PIL.Image.ANTIALIAS)
        
        # using BytesIO to save the resized image to memory
        with io.BytesIO() as output:
            img_resized.save(output, img.format)

            # if center is selected as left and top, picture is centered
            if left == "center" and top == "center":
                pic = self.slide.shapes.add_picture(output, 0, 0)
                pic.left = int((self.ss.slide_width - pic.width)/2) 
                pic.top = int((self.ss.slide_height - pic.height)/2) 
            
            # else position it depending on the left and top variable (in inches)
            else:
                self.slide.shapes.add_picture(output, Inches(left), Inches(top))

    # saves the pptx file
    def save(self, saveFile):
        self.ss.save(saveFile)
