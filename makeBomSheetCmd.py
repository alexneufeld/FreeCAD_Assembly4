#!/usr/bin/env python3
# coding: utf-8
# 
# makeBomCmd.py 
#
# BOM maker for ASM4
# creates a BOM as a FreeCAD spreadsheet

import os
import collections
from PySide import QtGui, QtCore
import FreeCADGui as Gui
import FreeCAD as App
from FreeCAD import Console as FCC
import pprint
import math

import libAsm4 as Asm4

def sortf(x):
  #print(x)
  if x == 'Name':
    return -2
  elif x == 'Quantity':
    return -1
  elif x == 'Misc. Info':
    return math.inf
  else:
    return ord(x[0][0].lower())-ord('a')

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
    # 'required' metadata set using a property of the model object
    metavalues = App.ActiveDocument.Model.RequiredPartMetadata
    # initiate the list to store metadata with the header line
    sheetdata = [["Assembly4 Bill of Materials"]]
    # get a list of parts in the model
    ObjInModel = App.ActiveDocument.Model.Group
    # strip out objects that aren't linked bodies or parts
    PartsList = []
    for obj in ObjInModel:
      if obj.TypeId == "App::Link":
        if obj.LinkedObject.TypeId in Asm4.containerTypes:
          PartsList.append(obj)
    # get unique parts and their frequency. the collections lib does it for us
    PartsCount = collections.Counter(PartsList)
    # pull all of the PartInfo properties assigned by the user into dictonaries
    datablock = []
    for thepart,num in PartsCount.items():
      lnkObj = thepart.LinkedObject
      subBlock = {"Quantity":num,"Name":thepart.LinkedObject.Label}
      for prop in lnkObj.PropertiesList:
        if lnkObj.getGroupOfProperty(prop) == 'PartInfo':
          subBlock.update({prop:getattr(lnkObj,prop)})
      datablock.append({k:v for k,v in sorted(subBlock.items(), key=sortf)})
    # properties that have assigned values for most of the parts (we set an 
    # arbitrary minimum of 50% for now) will get their own column in the
    # spreadsheet. properties that are only assigned to a few of the parts get
    # compressed into a single column at the end
    # at least X% of parts must have a property for it to get its own BOM column
    minDataCommonality = 0.5 
    keydump = [i for sl in list(map(lambda x: x.keys(),datablock)) for i in sl]
    keysCount = collections.Counter(keydump)
    commonKeys = []
    for key,ct in keysCount.items():
      if ct/len(PartsCount) >= minDataCommonality:
        commonKeys.append(key)
    commonKeys.sort(key=sortf)
    sheetdata.append(commonKeys+["Misc. Info"])
    for pdct in datablock:
      row = [""]*len(sheetdata[1])
      misc = ""
      for key,val in pdct.items():
        if key  not in commonKeys:
          misc += f"{key}: {val}, "
        else:
          for i,x in enumerate(commonKeys):
            if key == x:
              row[i] = val # this whole block is hacky ATM. fix it!
      row[-1] = misc
      sheetdata.append(row)
    # create the spreadsheet obbject
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
        # Adding/setting cells as generic properties causes WEIRD behaviour
        #theSheet.addProperty('App::PropertyString',cellname,"Base")
        #setattr(theSheet,cellname,str(cellval))
        theSheet.set(cellname,str(cellval))
    # add the page to the metadata group
    metagroup = App.ActiveDocument.getObject("Metadata")
    metagroup.addObject(theSheet)
    # final step - recompute
    App.activeDocument().recompute(None,True,True)
    # show the sheet to the user by selecting it:
    Gui.Selection.addSelection(App.ActiveDocument.Label,theSheet.Label)
    return

  

Gui.addCommand( 'Asm4_makeBOMSheet', makeBOMSheet())
