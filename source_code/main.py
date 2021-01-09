from pptx import Presentation
#For Picture
from pptx.util import Inches
#For Auotshapes
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE
#Change shape color
from pptx.enum.dml import MSO_THEME_COLOR
import  numpy as np, pandas as pd
#For excel
import os, copy
from time import perf_counter

t1 = perf_counter()

prs_input_path = r"C:\Users\adam.smith04\Documents\PROJECTS\mu_code\wavemark\resources\QBR template 2021.pptx"
prs_output_path = r"C:\Users\adam.smith04\Documents\PROJECTS\mu_code\wavemark\outputs\QBR template 2021.pptx"
df_file_path = r"C:\Users\adam.smith04\Documents\PROJECTS\mu_code\wavemark\resources\QBR Template FY21 v0.1.xlsx"

#Create data frame
df = pd.read_excel(df_file_path)

#creates a unique list of PAC RDs
idn_list = list(df['IDN_WaveMark'].unique())



prs = Presentation(prs_input_path)

#Add slides to a list
slides = list(prs.slides)

'''
for shape in slides[4].shapes:
    if shape.name == 'TextBox 6':
        #print(shape.name)
        #print(shape.shape_id)
'''

#give me the first customer in the list
for customer in idn_list[0:2]:
    #give me the 5th slide in the deck
    for slide in slides[4:5]:
        
        #give me the first shape in the group of shapes on the slide
        for shape in slide.shapes:
            if shape.name == 'TextBox 6':
                shape.text = customer 
                #to get the formatting of the text you have to 
                #add the text to the run through the text_frame
                text_frame = shape.text_frame
                p = text_frame.paragraphs[0]
                p.font.name = 'Arial (Body)'
                p.font.size = Pt(12)
                p.font.italic = True
                #set the text of the shape to customer
                #font: Arial (Body) 12 Italics
            

t2 = perf_counter()
print('Took {} Seconds'.format(t2-t1))
print('Saving powerpoint file...')
prs.save(prs_output_path)
os.startfile(prs_output_path)

    
