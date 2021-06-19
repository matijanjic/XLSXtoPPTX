# XLSXtoPPTX

A script that takes an existing excel file, reads the required columns that hold the info 
(word string, sentence string, word audio, sentence audio, sentence picture etc.) and arranges it all into a pptx presentation. The use case is pretty specific, my wife needed it for her phd experiment.
Audio files were generated using another script I made that uses the Google text-to-speech 
(gTTS) library that read the words and sentences from a csv file and exported them as .mp3's

## SlideShow<span></span>.py

Handles the powerpoint part.
I implemented a lot of methods that simplify adding empty slides, images (which can be resized and centered), text (also can be resized and centered), sound (that can autoplay) and it handles saving the pptx file.

*You have to take note that because it was made to have at most one sentence per slide, bigger text could possibly need adjusting.*

## XLSXtoPPTX<span></span>.py

This is where data is fed to the `SlideShow` class from the excel table. Here the slide layout, content and ordering is to be decided and for each row of the excel table a set of slides is made.

## Installation

Clone this repo using:
```bash
git clone https://github.com/matijanjic/XLSXtoPPTX.git
```
and install the requirements with:
```bash
pip install -r requirements.txt
```

## Usage

I really doubt that anybody will use this in this form, but better to be safe than sorry!

### SlideShow usage


#### Adding a slide
```python
addSlide()
```

#### Adding text
```python
addText(fontSize, text, width, height, left, top)
```
- fontSize - size of the font in points
- text - the text string
- width - width of the textbox in inches
- height - height of the textbox in inches
- left - position from the left edge in inches or 'centered' (default is centered)
- top - position from the top edge in inches or 'centered' (default is centered)

#### Adding an image
```python
addPicture(imgFile, maxSize, left, top)
```

- imgFile - image file location
- maxSize - maximum width or height of an image in pixels, no matter the image orientation
- left - same as in the addText() method
- top - same as in the addText() method

#### Adding sound
```python
addSound(soundFile)
```
- soundFile - sound file location

#### Autoplay

By default the sound is set to autoplay on slide show, but it is possible to turn that off by commenting out this part:

```python
tree = sound._element.getparent().getparent().getnext().getnext()
        timing = [el for el in tree.iterdescendants(
        ) if etree.QName(el).localname == 'cond'][0]
        timing.set('delay', '0')
```

#### Saving the pptx file
```python
save(saveFile)
```
- saveFile - save file location

### XLSXtoPPTX<span></span>.py

#### Sources
```python
def main():
    # some constants declared here
    xlsxFile = 'learning_pictures.xlsx'
    wordSoundFolder = 'sounds\words\\'
    sentenceSoundFolder = 'sounds\sentences\\'
    pictureFolder = 'images\pictures_learning\\'
```
These are the sources for the xlsx table, word and sentence sounds and images. One could add more and/or rename them if needed, but these ones worked for my use case.

Afther those ones come the column letters:

```python
wordTextCol = 'F'
wordSoundCol = 'G'
sentenceSoundCol = 'I'
sentencePictureCol = 'J'
colList = ['F', 'G', 'I', 'J']
```
which are the columns from which the data for the powerpoint will be extracted. 
`colList` is a list of all the needed columns listed in the above variables.

Here we define the start and end point for the rows and columns. 
```python
rowStart = 8
rowEnd = 282
colStart = 6
colEnd = 11
```
You could set it to the whole sheet, but this way you can save on time and memory. I could've automated this probably but didn't mind manually typing it in.

### Layout file

```python
slideShow = SlideShow('16x9.pptx', 6)
```
During the instancing of the `SlideShow` class, a layout file is required if you need a different aspect ratio than the default 4:3. A layout file is an empty pptx file that has the desired aspect ratio selected in the design tab. In the example above, that is a file named `16x9.pptx`

The `layoutNumber`, in this case 6, is the layout number from the python-pptx module. You can find more about it [here](https://python-pptx.readthedocs.io/en/latest/user/slides.html). It dictates which powerpoint layout is used.

### xlsxDict function call
```python
xlsxDict = getDictFromXlsx(xlsxFile, rowStart, rowEnd + 1, colStart, colEnd, colList)
```
This function returns a dictionary that holds the column letters and their values. The columns letters and their respectable variable names are passed as a list, this enables the program to scan more columns if needed.

### Main slide layout

This is where the final presentation is made. You can mix and match SlideShow methods to create a presentation you want.
```python
for i in range(rowEnd - rowStart + 1):
    slideShow.addSlide()
    slideShow.addText(80, 'X', 4, 1)
    slideShow.addSlide()
    slideShow.addText(44, xlsxDict[wordTextCol][i], 4, 1)
    slideShow.addSound(wordSoundFolder + xlsxDict[wordSoundCol][i])
    slideShow.addSlide()
    slideShow.addPicture(pictureFolder + xlsxDict[sentencePictureCol][i], 400)
    slideShow.addSound(sentenceSoundFolder + xlsxDict[sentenceSoundCol][i])
```
This code adds a slide with a big X in the middle in a 4x1 textbox. Then it adds another slide with a single word from the excel dictionary, after that it adds a sound and another slide with a picture and a yet another sound in it.

After that the presentation is saved and that's it.

### Conclusion

I'm aware this is a really narrow use case scenario program, but it's a project nontheless. I learned a lot and it's the first thing that got me out of tutorial hell. 

### Contributions 
If you'd wanted to contribute, don't hesitate to send a pull request or contact me via e-mail.