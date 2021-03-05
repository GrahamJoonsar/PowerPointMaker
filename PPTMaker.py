from pptx import Presentation
from PyDictionary import PyDictionary
from simple_image_download import simple_image_download as simp
from time import sleep

userInput = ''
topics = []
response = simp.simple_image_download
dictionary = PyDictionary()
PPTName = input("Name of PowerPoint: ")

print('Enter topics of slides: ')
while userInput != 'q':
    userInput = input('Topic of slide ')
    if userInput != 'q':
        topics.append([userInput, str(dictionary.meaning(userInput))])
        response().download(userInput , 1)


prs = Presentation()
for topic in topics:
    title_slide_layout = prs.slide_layouts[8]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    placeholder = slide.placeholders[1]
    pic = placeholder.insert_picture('simple_images/' + topic[0] + '/' + topic[0] + '_1' + '.jpeg')
    sub = slide.placeholders[2]
    title.text = topic[0].title()
    sub.text = topic[1]

prs.save(PPTName + '.pptx')