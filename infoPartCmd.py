#!/usr/bin/env python3
# coding: utf-8
#
# LGPL
# Copyright HUBERT Zoltán
#
# infoPartCmd.py



import os
import re
from PySide import QtGui, QtCore
import FreeCADGui as Gui
import FreeCAD as App
from FreeCAD import Console as FCC


import libAsm4 as Asm4



"""
    +-----------------------------------------------+
    |               Helper functions                |
    +-----------------------------------------------+
"""

# allowed types to edit info
partTypes = [ 'App::Part', 'PartDesign::Body']

def checkPart():
    selectedPart = None
    # if an App::Part is selected
    if len(Gui.Selection.getSelection())==1:
        selectedObj = Gui.Selection.getSelection()[0]
        if selectedObj.TypeId in partTypes:
            selectedPart = selectedObj
    return selectedPart



"""
    +-----------------------------------------------+
    |                  The command                  |
    +-----------------------------------------------+
"""
class infoPartCmd():
    def __init__(self):
        super(infoPartCmd,self).__init__()

    def GetResources(self):
        return {"MenuText": "Edit Part Information",
                "ToolTip": "Edit Part Information",
                "Pixmap" : os.path.join( Asm4.iconPath , 'Asm4_PartInfo.svg')
                }

    def IsActive(self):
        # We only insert a link into an Asm4  Model
        if App.ActiveDocument and checkPart():
            return True
        return False

    def Activated(self):
        Gui.Control.showDialog( infoPartUI() )




"""
    +-----------------------------------------------+
    |    The UI and functions in the Task panel     |
    +-----------------------------------------------+
"""
class infoPartUI():

    def __init__(self):
        self.base = QtGui.QWidget()
        self.form = self.base        
        iconFile = os.path.join( Asm4.iconPath , 'Asm4_PartInfo.svg')
        self.form.setWindowIcon(QtGui.QIcon( iconFile ))
        self.form.setWindowTitle("Edit Part Information")
       
        # hey-ho, let's go
        self.part = checkPart()
        self.reqPartInfo = App.ActiveDocument.Model.RequiredPartMetadata
        self.makePartInfo()
        self.infoTable = []
        self.getPartInfo()
        self.userinput = {}
        self.checkedProps = {}

        # the GUI objects are defined later down
        self.drawUI()


    def getPartInfo(self):
        for prop in self.part.PropertiesList:
            if self.part.getGroupOfProperty(prop)=='PartInfo' :
                if self.part.getTypeIdOfProperty(prop)=='App::PropertyString' :
                    value = self.part.getPropertyByName(prop)
                    self.infoTable.append([prop,value])

    def makePartInfo( self, reset=False ):
        # add the default part information
        # required metadata fields are currently specified by a parameter in 
        # the model object. I don't know if this is the best possible solution.
        # They could also eg: be set using global ASM4 preferences 
        for info in self.reqPartInfo:
            if not hasattr(self.part,info):
                self.part.addProperty( 'App::PropertyString', info, 'PartInfo' )
        return

    # close
    def finish(self):
        Gui.Control.closeDialog()

    # standard panel UI buttons
    def getStandardButtons(self):
        return int(QtGui.QDialogButtonBox.Cancel | QtGui.QDialogButtonBox.Ok)

    # Cancel
    def reject(self):
        self.finish()

    # OK: we insert the selected part
    def accept(self):
        for prop in self.infoTable:
            #prop[1] = str()
            prop[1] = self.userinput[prop[0]].text()
            setattr(self.part,prop[0],prop[1])
        self.finish()


    # Define the iUI, only static elements
    def drawUI(self):
        # Place the widgets with layouts
        self.mainLayout = QtGui.QVBoxLayout(self.form)
        self.formLayout = QtGui.QFormLayout()

        for prop in self.infoTable:
            checkLayout = QtGui.QHBoxLayout()
            propValue   = QtGui.QLineEdit()
            propValue.setText( str(prop[1]) )
            checkLayout.addWidget(propValue)
            # store the text boxes in a dictionary so we can access them later
            self.userinput.update({prop[0]: propValue})
            # disable delete checkboxes for default properties
            if prop[0] not in self.reqPartInfo:
                checked     = QtGui.QCheckBox()
                self.checkedProps.update({prop[0]: checked})
                checkLayout.addWidget(checked)
            # add spaces to camelCased variable names for readability
            rowLabel = re.sub(r"\B([A-Z])", r" \1", prop[0])
            self.formLayout.addRow(QtGui.QLabel(rowLabel),checkLayout)

        self.mainLayout.addLayout(self.formLayout)
        self.mainLayout.addWidget(QtGui.QLabel())
        
        # Buttons
        self.buttonsLayout = QtGui.QHBoxLayout()
        self.AddNew = QtGui.QPushButton('Add New Info')
        self.Delete = QtGui.QPushButton('Delete Selected')
        self.buttonsLayout.addWidget(self.AddNew)
        self.AddNew.clicked.connect(self.newProp)
        self.buttonsLayout.addStretch()
        self.buttonsLayout.addWidget(self.Delete)
        self.Delete.clicked.connect(self.removeProp)

        self.mainLayout.addLayout(self.buttonsLayout)
        self.form.setLayout(self.mainLayout)

        # Actions
    def newProp(self):
        text,ok = QtGui.QInputDialog.getText(None, "Create a New Property", 'Enter new Property name :'+' '*30, text = "")
        if text and ok:
            atext = "".join(text.split())
            self.part.addProperty( 'App::PropertyString', atext, 'PartInfo' )
            self.infoTable.append([atext, ""])
            # The only way I could get the removed/added properties to appear and 
            # disappear correctly was to close and relaunch the entire widget
            Gui.Control.closeDialog()
            Gui.runCommand('Asm4_infoPart',0)

    def removeProp(self):
        removedPropNames = []
        for prop in self.infoTable:
            if prop[0] in self.checkedProps:
                if self.checkedProps[prop[0]].isChecked():
                    self.part.removeProperty(prop[0])
                    removedPropNames.append(prop[0])
        self.infoTable = [item for item in self.infoTable if item[0] not in removedPropNames]
        self.checkedProps = {}
        Gui.Control.closeDialog()
        Gui.runCommand('Asm4_infoPart',0)






"""
    +-----------------------------------------------+
    |       add the command to the workbench        |
    +-----------------------------------------------+
"""
Gui.addCommand( 'Asm4_infoPart', infoPartCmd() )

