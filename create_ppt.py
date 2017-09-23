# -*- coding: utf-8 -*-
"""
#   -------------------------------------------------   #
#
#   -------- Description --------
#   1) read.\template.xlsx  as Pandas Dataframe und execute each
#   line in the template.xlsx (each Befehl is one function)
#   2) read template.pptx
#   3) write new_output.pptx with commands from template.xlsx
#   -------- Run --------
#   creat_ppt.py
#
#
#   Created on Fri Sep 22 20:04:13 2017 by   mom
#   -------------------------------------------------   #
"""
import pandas as pd
import os, re
from pptx import Presentation
#from pptx.util import Inches

xls_template = 'template.xlsx'
ppt_template = 'template.pptx'
pptNew = 'new_output.pptx'

global actslide

# --------------------------------------------------------------------
#            Definition der Funktione
# --------------------------------------------------------------------

def newPage(template=0):
    """
    Describtion
    """
    title_slide_layout = prs.slide_layouts[int(template)]
    actslide = prs.slides.add_slide(title_slide_layout)
    global actslide

def tiff(placeholder=1, pfad=''):
    """
    Describtion
    """
    assert os.path.isfile(pfad), 'File {} does not excist'.format(pfad)
    placeholder = actslide.placeholders[int(placeholder)]
    placeholder.insert_picture(pfad)

def write(placeholder=1, text=''):
    """
    Describtion
    """
    subtitle = actslide.placeholders[int(placeholder)]
    subtitle.text = text

def title(text=''):
    """
    Describtion
    """
    title = actslide.shapes.title
    title.text = text


functionDict = {'newpage': newPage,
                'tiff': tiff,
                'write': write,
                'title': title}
# ------------------------------------------------------------



# --------------------------------------------------------------------
#            Öffne PPT und lies die Excel
# --------------------------------------------------------------------

prs = Presentation(ppt_template)

xls = pd.read_excel(xls_template, index='Befehl')
xls = xls[~xls.Befehl.str.contains(r'^#')]    # drop comment lines

for index, actBefehl in xls.iterrows():
    # --- füge values als ParameterDict zusammen
    actParameterList = actBefehl.values[1:]   #  alle spalten nach Befehl\
                                          #    array['Template(1)' nan]
    Parameter = {}
    for actParameter in actParameterList:
        try:
            match = re.findall(r'(\w*)\(\ *(.*)\ *\)', actParameter)
            Parameter.update({match[0][0].lower():match[0][1]})
        except IndexError:
            pass
        except TypeError:
            pass

    # --- run function
    actBefehl = actBefehl.Befehl.replace(' ','')   # delete space
    #try:
    functionDict[actBefehl.lower()](**Parameter)
    #except KeyError:
    #    print ('Warnung, Befehl nicht gefunden. Bitte Überprüfen')
    #    print ('Zeile {}  Text: "{}"'.format(index, actBefehl))

prs.save(pptNew)