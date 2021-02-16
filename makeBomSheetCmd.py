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
import libAsm4 as Asm4
import FastenersLib
import pprint
# helper function
def sortf(x):
  if x == 'Name':
    return -2
  elif x == 'Quantity':
    return -1
  elif x == 'Misc. Info':
    return 999
  else:
    return ord(x[0][0].lower())-ord('a')

# this function converts a 0-based numerical index to the lettered column 
# indexes used by spreadsheet software. works up to x=701
def index(x):
  return (x>25)*chr(int(((x-26)-(x-26)%26)/26+65))+chr(int(x%26+65))

class makeBOMSheet:

  def __init__(self):
    super(makeBOMSheet,self).__init__()
    self.UI = QtGui.QDialog()
    self.selectionUI()

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
    if Asm4.checkModel():
      return True
    else:
      return False

  def Activated(self):
    self.UI.show()

  def selectionUI(self):
    self.UI.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
    self.UI.setWindowTitle('Select BOM Type')
    self.UI.setWindowIcon(QtGui.QIcon( os.path.join(Asm4.iconPath,'Asm4_BOM_Sheet.svg')))
    self.UI.setMinimumWidth(300)
    self.UI.resize(300,110)
    self.UI.setModal(False)
    # the layout for the main window is vertical (top to down)
    self.mainLayout = QtGui.QVBoxLayout(self.UI)
    # Define the fields for the form ( label + widget )
    self.formLayout = QtGui.QFormLayout()
    self.bomselector = QtGui.QComboBox()
    for s in ["Parts","Fasteners"]:
      self.bomselector.addItem(s)
    self.formLayout.addRow(QtGui.QLabel('Type'),self.bomselector)
    # apply the layout
    self.mainLayout.addLayout(self.formLayout)
    self.mainLayout.addStretch()
    # Buttons
    self.buttonLayout = QtGui.QHBoxLayout()
    self.CancelButton = QtGui.QPushButton('Cancel')
    self.OkButton = QtGui.QPushButton('OK')
    self.OkButton.setDefault(True)
    # the button layout
    self.buttonLayout.addWidget(self.CancelButton)
    self.buttonLayout.addStretch()
    self.buttonLayout.addWidget(self.OkButton)
    self.mainLayout.addLayout(self.buttonLayout)
    # finally, apply the layout to the main window
    self.UI.setLayout(self.mainLayout)
    # Actions
    self.CancelButton.clicked.connect(self.onCancel)
    self.OkButton.clicked.connect(self.onOK)

  def onCancel(self):
    self.UI.close()

  def onOK(self):
    Bom_Type = self.bomselector.currentText()
    if Bom_Type == "Parts":
      self.makePartsBOM()
    elif Bom_Type == "Fasteners":
      self.makeFastenersBOM()
    self.UI.close()

  def makePartsBOM(self):
    # initiate the list to store metadata with the header line
    sheetData = [["Assembly4 Bill of Materials"]]
    # get a list of parts in the model, and
    # strip out objects that aren't linked bodies or parts
    PartsList = []
    for obj in Asm4.checkModel().Group:
      if Asm4.isLinkToPart(obj):
        PartsList.append(obj)
    # get unique parts and their frequency. the collections lib does it for us
    PartsCount = collections.Counter(PartsList)
    # pull all of the PartInfo properties assigned by the user into dictionaries
    datablock = []
    for thepart,num in PartsCount.items():
      lnkObj = thepart.LinkedObject
      subBlock = {"Quantity":num,"Name":thepart.LinkedObject.Label}
      for prop in lnkObj.PropertiesList:
        if lnkObj.getGroupOfProperty(prop) == 'PartInfo':
          subBlock.update({prop:getattr(lnkObj,prop)})
      datablock.append(subBlock)
    for d in datablock:
      print(d)
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
    #commonKeys = ["Name","Quantity"]+commonKeys.remove()
    sheetData.append(commonKeys+["Misc. Info"])
    for pdct in datablock:
      row = [""]*len(sheetData[1])
      misc = ""
      for key,val in pdct.items():
        if key  not in commonKeys:
          misc += f"{key}: {val}, "
        else:
          for i,x in enumerate(commonKeys):
            if key == x:
              row[i] = val # this whole block is hacky ATM. fix it!
      row[-1] = misc
      sheetData.append(row)
    # create the spreadsheet object
    theSheet = App.activeDocument().addObject('Spreadsheet::Sheet','ASM4_BOM')
    # format the title by merging cells and setting background colours
    titlerange = 'A1:'+index(len(sheetData[-1])-1)+'1'
    theSheet.mergeCells(titlerange)
    # titles have yellow and purple backgrounds to match ASM4 colorscheme
    theSheet.setBackground(titlerange, (0.666667,0.666667,1.000000,1.000000))
    theSheet.setStyle(titlerange, 'bold')
    theSheet.setAlignment(titlerange, 'center|vcenter|vimplied')
    # repeat for column subtitles
    for col in range(len(sheetData[-1])):
      thecell = index(col)+'2'
      theSheet.setBackground(thecell, (1.000000,0.996078,0.792157,1.000000))
    # cram the data into the correct cells
    for i,row in enumerate(sheetData):
      for j,cellval in enumerate(row):
        theSheet.set(index(j)+str(i+1),str(cellval))
    # add the page to the metadata group
    metagroup = App.ActiveDocument.getObject("Metadata")
    metagroup.addObject(theSheet)
    # final steps - recompute...
    App.activeDocument().recompute(None,True,True)
    # and show the sheet to the user by selecting it:
    Gui.Selection.addSelection(App.ActiveDocument.Label,theSheet.Label)
    return

  def makeFastenersBOM(self):
    # check that the FastenersWB is installed and ready-to-go
    if 'FastenersWorkbench' in Gui.listWorkbenches():
      import FastenersCmd
      import ScrewMaker
    else:
      print("Fasteners WorkBench not found\nCould not create Fasteners BOM")
      return
    # collect the required data from the Model
    FSDict = {}
    for obj in App.ActiveDocument.Model.Group:
      if FastenersLib.isFastener(obj):
        FSClass = ScrewMaker.screwTables[obj.type][0]
        if FSClass == "Screw":
          if obj.length != "Custom":
            key = obj.diameter + "x" + obj.length + " Screw" + " - " + obj.type
            thelength = obj.length
          else:
            key = obj.diameter + "x" + obj.lengthCustom + " Screw" + " - " + obj.type
            thelength = obj.lengthCustom
        else:
          key = obj.diameter + " " + FSClass + " - " + obj.type
          thelength = ""
        if key not in FSDict:
          FSDict.update({key:{"qty":1,"standard":obj.type,"diameter":obj.diameter,"length":thelength,"class":FSClass}})
        else:
          FSDict[key].update({"qty":FSDict[key]["qty"]+1})
    FSDict = dict(sorted(FSDict.items()))
    pprint.pprint(FSDict)
    # compile the data to a nicely formatted spreadsheet
    theSheet = App.activeDocument().addObject('Spreadsheet::Sheet','ASM4_FS_BOM')
    titlerange = 'A1:C1'
    theSheet.mergeCells(titlerange)
    # titles have yellow and purple backgrounds to match ASM4 colorscheme
    theSheet.setBackground(titlerange, (1.0000,0.8314,0.5000,1.0000))
    theSheet.setStyle(titlerange, 'bold')
    theSheet.setAlignment(titlerange, 'center|vcenter|vimplied')
    theSheet.set("A1","Assembly4 Fasteners BOM")
    # column subtitles
    for n, subtitle in enumerate(["Quantity","Description","Standard"]):
      thecell = index(n)+'2'
      theSheet.setBackground(thecell, (0.5747,0.8275,0.7922,1.0000))
      theSheet.set(thecell,subtitle)
    # cram the data into the correct cells
    for row, (key,valdict) in enumerate(FSDict.items()):
      theSheet.set("A"+str(row+3),str(valdict["qty"]))
      if valdict["class"] == "Screw":
        descstr = valdict["diameter"] + "x" + valdict["length"] + " Screw"
      else:
        descstr = valdict["diameter"] + " " + valdict["class"]
      theSheet.set("B"+str(row+3),descstr)
      theSheet.set("C"+str(row+3),valdict["standard"])
    # add the page to the metadata group
    metagroup = App.ActiveDocument.getObject("Metadata")
    metagroup.addObject(theSheet)
    # final steps - recompute...
    App.activeDocument().recompute(None,True,True)
    # and show the sheet to the user by selecting it:
    Gui.Selection.addSelection(App.ActiveDocument.Label,theSheet.Label)


Gui.addCommand( 'Asm4_makeBOMSheet', makeBOMSheet())
