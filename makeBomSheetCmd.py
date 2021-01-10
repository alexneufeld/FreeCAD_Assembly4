#!/usr/bin/env python3
# coding: utf-8
# 
# makeBomCmd.py 
#
# BOM maker for ASM4
# creates a BOM as a FreeCAD spreadsheet

import os

from PySide import QtGui, QtCore
import FreeCADGui as Gui
import FreeCAD as App
from FreeCAD import Console as FCC

import libAsm4 as Asm4

class makeBOMSheet:
  def __init__(self):
    pass

  def GetResources(self):
    return {"MenuText": "Create Part List to spreadsheet",
            "ToolTip": "Create the Bom (Bill of Materials) of an Assembly4 Model",
            "Pixmap" : os.path.join( Asm4.iconPath , 'Asm4_BOM_Sheet.svg')
            }

  def checkModel(self):
    # check whether there is already a Model in the document
    # Returns True if there is an object called 'Model'
    if App.ActiveDocument and App.ActiveDocument.getObject('Model') and App.ActiveDocument.Model.TypeId=='App::Part':
      return(True)
    else:
      return(False)

  def IsActive(self):
    return self.checkModel()

  def Activated(self):
    # we want to compile the data in a format that python handles nicely,
    # then translate it to a FreeCAD sheet as a final step -> makes it 
    # easier to add other export formats eg .ods, .csv
    # attempt one: recursive list, sublists are rows
    metavalues = App.ActiveDocument.Model.RequiredPartMetadata
    sheetdata = [["Assembly4 Bill of Materials"]]
    sheetdata.append(["Part Name","Quantity"]+metavalues)
    # get a list of parts in the model
    ObjInModel = App.ActiveDocument.Model.Group
    PartsList = []
    for obj in ObjInModel:
      # we only care about linked bodies and parts
      if obj.TypeId == "App::Link":
        if obj.LinkedObject.TypeId in Asm4.containerTypes:
          PartsList.append(obj)
    # get unique parts and their frequency
    import collections
    PartsCount = collections.Counter(PartsList)
    for thepart,num in PartsCount.items():
      sheetrow = [thepart.LinkedObject.Label,num]
      for metakey in metavalues:
        sheetrow.append(getattr(thepart,metakey))
      sheetdata.append(sheetrow)
    # the BOM list should now be completed!
    # now we can spit out a FreeCAD spreadsheet
    theSheet = App.activeDocument().addObject('Spreadsheet::Sheet','ASM4_BOM')
    # this lambda converts a 0-based numerical index to the lettered column 
    # indexes used by spreadsheet software. works up to x=701
    index = lambda x: (x>25)*chr(int(((x-26)-(x-26)%26)/26+65))+chr(int(x%26+65))
    # format the title by merging cells and setting background colours
    # we must do this BEFORE inputting data, as changing cell formats
    # blanks out the data in them for some reason :P
    titlerange = 'A1:'+index(len(sheetdata[-1])-1)+'1'
    theSheet.mergeCells(titlerange)
    # titles have yellow and purple backgrounds to match ASM4 colorscheme
    theSheet.setBackground(titlerange, (0.666667,0.666667,1.000000,1.000000))
    theSheet.setStyle(titlerange, 'bold')
    theSheet.setAlignment(titlerange, 'center|vcenter|vimplied')
    # repeat for column subtitles
    for col in range(len(sheetdata[-1])):
      thecell = index(col)+'2'
      theSheet.setBackground(thecell, (1.000000,0.996078,0.792157,1.000000))
    App.activeDocument().recompute(None,True,True)
    # cram the data into the correct cells
    for i,row in enumerate(sheetdata):
      for j,cellval in enumerate(row):
        cellname = index(j)+str(i+1)
        theSheet.addProperty('App::PropertyString',cellname,"Base")
        setattr(theSheet,cellname,str(cellval))
    # add the page to the metadata group
    metagroup = App.ActiveDocument.getObject("Metadata")
    metagroup.addObject(theSheet)
    # final step - recompute
    App.activeDocument().recompute(None,True,True)
    # show the sheet to the user by selecting it:
    Gui.Selection.addSelection(App.ActiveDocument.Label,theSheet.Label)
    return
    
Gui.addCommand( 'Asm4_makeBOMSheet', makeBOMSheet())
