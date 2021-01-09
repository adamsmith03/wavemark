from pptx import Presentation
#For Picture
from pptx.util import Inches
#For Auotshapes
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
df_file_path = r"C:\Users\adam.smith04\Documents\PROJECTS\mu_code\wavemark\resources\QBR Template FY21.xlsx"

prs = Presentation(prs_input_path)
print(prs)