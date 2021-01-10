#!/usr/bin/env python3
# coding: utf-8
# 
# makeDrawingCmd.py 
#
# This command automagically creates a TechDraw drawing, in a (hopefully) 
# sensible format, of the current documents assembly4 model. The idea is to 
# automate the creation of good quality assembly metadata as much as possible.
#

import os

from PySide import QtGui, QtCore
import FreeCADGui as Gui
import FreeCAD as App
from FreeCAD import Console as FCC
from FreeCAD import Base
import math
import libAsm4 as Asm4
from datetime import date

class makeDWG:
    def __init__(self):
        pass

    def checkModel(self):
        # check whether there is already a Model in the document
        # Returns True if there is an object called 'Model'
        if App.ActiveDocument and App.ActiveDocument.getObject('Model') and App.ActiveDocument.Model.TypeId=='App::Part':
            return(True)
        else:
            return(False)
    
    def IsActive(self):
        return self.checkModel()

    def GetResources(self):
        return {"MenuText": "Create a Drawing of the Model",
                "ToolTip": "Create a TechDraw Drawing of an Assembly4 Model",
                "Pixmap" : os.path.join( Asm4.iconPath , 'Asm4_Dwg.svg')
                }

    def Activated(self):
        # create the drawing object:
        drawpage = App.activeDocument().addObject('TechDraw::DrawPage','Page')
        drawtemplate = App.activeDocument().addObject('TechDraw::DrawSVGTemplate','Template')
        # most of the techdraw default templates aren't exactly pretty
        # the following are the least bad IMO:
        # A4_Landscape_ISO7200_Pep.svg
        # A4_Landscape_ISO7200TD.svg
        # FreeCAD comes with these templates preinstalled
        # But their location shanges based on the user's platform
        # TODO: implement logic to find the template location
        # for now, they are mirrored in the ASM4 resources folder
        # we could also design some nicer templates for ASM4
        templatesPath = os.path.join( Asm4.wbPath, 'Resources/dwg_templates' )
        thetemplate = "A4_Landscape_ISO7200TD.svg"
        drawtemplate.Template = str(os.path.join(templatesPath,thetemplate))
        drawpage.Template = drawtemplate
        # that's basic setup done.
        # now we create views of the model object on the page
        # this is a basic isometric view
        isoview = App.activeDocument().addObject('TechDraw::DrawViewPart','View')
        drawpage.addView(isoview)
        themodel = App.ActiveDocument.getObject("Model")
        isoview.Source = themodel
        isoview.Direction = Base.Vector(0.577,-0.577,0.577)
        isoview.XDirection = Base.Vector(0.707,0.707,0.000)
        isoview.ScaleType = "Page"
        # we will set the scale of the view so that It fills a reasonable
        # portion of the page. Fortunately, we are able to access the size
        # of the model using themodel.ViewObject.getBoundingBox()
        bbox = themodel.ViewObject.getBoundingBox()
        modelsize = bbox.DiagonalLength
        pagesize = math.sqrt(drawtemplate.Width.Value**2+drawtemplate.Height.Value**2)
        # set the page scale correctly
        # 8 is a scale factor. we want the view to take up about 1/5 of the page area
        # (x * modelsize)/pagesize = 1/5
        # note that the math here is totally estimative, no hard numbers
        # it will need to be improved eventually
        x = pagesize/(5*modelsize)
        pscale = str(round(math.log2(x))) if x>=1 else "1/"+str(round(1/x))
        drawpage.setExpression("Scale",pscale)
        # the centroid of the view object is very roughly the center of its bounding box
        # we will try to put the view in a sensible position on the page using the info 
        # we have available to us:
        # 10 and 20 are placeholders for the margins of the page
        # this part of the code needs to be smarter!
        isoview.X = str(10+0.5*modelsize)+' mm'
        isoview.Y = str(drawtemplate.Height.Value-20-0.5*modelsize)+' mm'
        # now let's make a 3 view viewset and put it on the page:
        projgroup = App.activeDocument().addObject('TechDraw::DrawProjGroup','ProjGroup')
        drawpage.addView(projgroup)
        projgroup.Source = themodel
        projgroup.addProjection('Front')
        projgroup.addProjection('Top')
        projgroup.addProjection('Right')
        projgroup.Anchor.Direction = Base.Vector(0.000,-1.000,0.000)
        projgroup.Anchor.RotationVector = Base.Vector(1.000,0.000,0.000)
        projgroup.Anchor.XDirection = Base.Vector(1.000,0.000,0.000)
        projgroup.ScaleType = "Page"
        projgroup.ProjectionType = "Third Angle"
        # we will leave the 3-view in the center of the page, where it is placed by default
        # add the page to the metadate group
        metagroup = App.ActiveDocument.getObject("Metadata")
        metagroup.addObject(drawpage)
        # the editable text fields are stored as follows:
        fields = {}
        # another dict to handle filling text fields automagically
        fieldhandler = {
            "Designed_by_Name": App.ActiveDocument.CreatedBy,
            "AUTHOR_NAME": App.ActiveDocument.CreatedBy,
            "OWNER_NAME": App.ActiveDocument.CreatedBy,
            "FC-DATE": str(date.today()),
            "DATE": str(date.today()),
            "FC-SC": pscale,
            "Scale": pscale,
            "SCALE": pscale,
            "FC-SH": "1 of 1",
            "Sheet": "1 of 1",
            "FC-Title": App.ActiveDocument.Label,
            "Title": App.ActiveDocument.Label,
            "DRAWING_TITLE": App.ActiveDocument.Label,
            "TITLELINE-1": App.ActiveDocument.Label,
            "Subtitle": App.ActiveDocument.FileName,
            "TITLELINE-2": App.ActiveDocument.FileName,
            "SI-1": App.ActiveDocument.FileName,
        }
        for key in drawtemplate.EditableTexts:
            try:
                fields.update({key: fieldhandler[key]})
            except KeyError:
                fields.update({key: drawtemplate.EditableTexts[key]})
        # FreeCAD only accepts the change to the editabletexts property
        # if we reassign the whole dictionary.
        # EditableTexts.update({k:v}) and EditableTexts[k] = v do nothing 
        drawtemplate.EditableTexts = fields
        # techdraw likes recomputes
        App.activeDocument().recompute(None,True,True)
        # show the user the dwg by selecting it
        Gui.Selection.addSelection(App.ActiveDocument.Label,drawpage.Label)

Gui.addCommand( 'Asm4_makeDWG', makeDWG() )

# notes
'''

A4_LandscapeTD.svg
{'Designed_by_Name': 'Designed by Name', 'Drawing_number': 'Drawing number', 'FC-Date': 'Date', 
'FC-SC': 'Scale', 'FC-SH': 'Sheet', 'FC-Title': 'Title', 'Subtitle': 'Subtitle', 'Weight': 'Weight'}

A4_Landscape_ISO7200TD.svg
{'AUTHOR_NAME': 'AUTHOR NAME', 'DN': 'DN', 'DRAWING_TITLE': 'DRAWING TITLE', 'FC-DATE': 'DD/MM/YYYY',
 'FC-REV': 'REV A', 'FC-SC': 'SCALE', 'FC-SH': 'X / Y', 'FC-SI': 'A4', 'FreeCAD_DRAWING': 'FreeCAD DRAWING',
  'PN': 'PN', 'SI-1': '', 'SI-3': ''}

A4_Landscape_ISO7200_Pep.svg
{'APPROVER_NAME': 'APPROVER NAME', 'AUTHOR_NAME': 'AUTHOR NAME', 'DATE': 'YYYY-MM-DD',
 'DN': 'DN', 'DOCUMENT_TYPE': '', 'OWNER_NAME': 'OWNER NAME', 'PM': 'PM', 'PN': 'PN',
  'REVISION': 'REV A', 'RIGHTS': "(R) DO NOT DUPLICATE THIS DRAWING TO THIRD PARTIES WITHOUT OWNER'S PERMISSION !",
   'SCALE': 'M x:x', 'SHEET': '99 of 99', 'SIZE': 'A4', 'TITLELINE-1': 'FreeCAD', 'TITLELINE-2': ''
   , 'TITLELINE-3': '', 'TOLERANCE': '+/- ?'}

'''