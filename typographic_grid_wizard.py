#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
VERSION: 1.2 of 2019-11-06 (works for me with Scribus 1.5.5 and python 2.7 on Windows 10).
AUTHOR: Rafferty River. 
LICENSE: GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007. 
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY.

DESCRIPTION & USAGE:
Generates a grid of guides where the cells follow the baseline grid. You can use this script on a master page and then apply it to other pages). 
Input parameters are: margins and gutters in baseline units, number of columns and rows and desired leading in points. 

Then top and bottom margins are adapted with trial and error until the result fits. Then you are proposed to draw the guides. 

Remark: the initial values below are used as an example for an A4 document.  You can change them to set your own defaults.
"""
##################################################
# imports

import sys

try:
    import scribus
    from scribus import UNIT_POINTS
except ImportError:
    print('This Python script is written for the Scribus scripting interface.')
    print('It can only be run from within Scribus.')
    sys.exit(1)

try:
    if sys.version_info[0] <= 2: # for python 2 or below
        import Tkinter as tk
        import tkMessageBox
    else: # for python 3
        import tkinter as tk
        import tkinter.messagebox as tkMessageBox 
except ImportError:
    print("This script requires Python's Tkinter properly installed.")
    messageBox('Script failed',
               'This script requires Python\'s Tkinter properly installed.',
               ICON_CRITICAL)
    sys.exit(1)

##################################################
# calculate and draw the typographic grid

def calculateGrid(removeGuides, lineHeightPt, xHeightPt, marginTop, marginBottom, marginLeft, marginRight, rowCount, columnCount, hGutter, vGutter):
    """ """
    if (columnCount == 0 or rowCount == 0 or lineHeightPt == 0):
        tkMessageBox.showwarning("Error:","  Number of columns, number of rows,\n  or leading cannot be 0!\n  Please try again.")
        return

    if (marginTop<0 or marginBottom<0 or marginLeft<0 or marginRight<0 or columnCount<0 or rowCount<0 or vGutter<0 or hGutter<0 or lineHeightPt<0):
        tkMessageBox.showwarning("Error:","  Negative entries are not allowed!\n  Please try again.")
        return

    pageWidth, pageHeight = scribus.getPageSize()

    pageLineCount = pageHeight / lineHeightPt

    rowLines = (pageLineCount - marginTop - marginBottom - hGutter * (rowCount - 1)) / rowCount
    if ((rowLines - round(rowLines)) * rowCount + marginTop + marginBottom) < -0.01 : # to avoid negative overflow corrections greater than the top+bottom margin lines
        rowLines = int(rowLines)
    else: # find the smallest overflow correction
        rowLines = round(rowLines)

    hOverflow = pageLineCount - marginTop - marginBottom - rowLines * rowCount - hGutter * (rowCount - 1)
    changeLineHeightPt = lineHeightPt * pageLineCount / (marginTop + marginBottom + rowLines * rowCount + hGutter * (rowCount - 1))

    columnUnitPt = lineHeightPt

    columnUnitsCount = pageWidth / columnUnitPt
    columnLines = (columnUnitsCount - marginLeft - marginRight - vGutter * (columnCount - 1)) / columnCount

    if not (abs(hOverflow) < 0.01):
        tkMessageBox.showwarning("Please try again:"," Change leading to "+str(round(changeLineHeightPt,3))+" pt\n or change top/bottom margins by "+str(round(hOverflow,2))+" lines \n or change number of rows")

    else: # all entries fit the calculation
        cellHeight = lineHeightPt * rowLines
        hGuides = []
        previousGuide = marginTop*lineHeightPt
        if not xHeightPt == 0.0:
           hGuides.append(previousGuide + (lineHeightPt - xHeightPt))
        for i in range(1, rowCount):
            hGuides.append(previousGuide + cellHeight)
            hGuides.append(previousGuide + cellHeight + hGutter*lineHeightPt)
            if not xHeightPt == 0.0:
                hGuides.append(previousGuide + cellHeight + hGutter*lineHeightPt + (lineHeightPt - xHeightPt))
            previousGuide = previousGuide + cellHeight + hGutter*lineHeightPt

        cellWidth = columnUnitPt * columnLines
        vGuides = []
        previousGuide = marginLeft*columnUnitPt
        for i in range(1, columnCount):
            vGuides.append(previousGuide + cellWidth)
            vGuides.append(previousGuide + cellWidth + vGutter*columnUnitPt)
            previousGuide = previousGuide + cellWidth + vGutter*columnUnitPt

        # draw the typographic grid

        msg=tkMessageBox.askokcancel("INFO:", " CALCULATION RESULT:\n line height="+str(lineHeightPt)+" pt\n cell height="+str(cellHeight)+" pt\n cell width="+str(cellWidth)+" pt\n cell ratio (h/w)="+str(cellHeight/cellWidth)+"\n\n APPLY GUIDES?\n (guides will be drawn after you Quit)")
        if msg == True:
            scribus.setBaseLine(lineHeightPt, lineHeightPt*(marginTop-int(marginTop)))
            scribus.setMargins(marginLeft*columnUnitPt, marginRight*columnUnitPt, marginTop*lineHeightPt, marginBottom*lineHeightPt) ### scribus.setPageMargins() does not yet exist in the API ###

            if removeGuides == True:
                scribus.setHGuides(hGuides)
                scribus.setVGuides(vGuides)
            else:
                scribus.setHGuides(scribus.getHGuides()+hGuides)
                scribus.setVGuides(scribus.getVGuides()+vGuides)

##################################################
# create gui with Tkinter

class TkGrid(tk.Frame):
    def __init__(self, master=None):
        """ Setup the dialog """
        self.key = 'default'
        tk.Frame.__init__(self, master)
        self.grid()
        self.master.title('Typographic Grid Wizard')
        
        #define variables
        self.removeGuides = tk.BooleanVar()
        self.lineHeightPt = tk.DoubleVar()
        self.xHeightPt = tk.DoubleVar()
        self.marginTop = tk.DoubleVar()
        self.marginBottom = tk.DoubleVar()
        self.marginLeft= tk.DoubleVar()
        self.marginRight = tk.DoubleVar()
        self.rowCount = tk.IntVar()
        self.columnCount = tk.IntVar()
        self.hGutter = tk.IntVar()
        self.vGutter = tk.DoubleVar()

        # initial values (you can change them into your own defaults!)
        self.removeGuides.set(True)
        self.lineHeightPt.set(14.77)
        self.xHeightPt.set(0.0)
        self.marginTop.set(2.0)
        self.marginBottom.set(2.0)
        self.marginLeft.set(2.0)
        self.marginRight.set(2.0)
        self.rowCount.set(8)
        self.columnCount.set(6)
        self.hGutter.set(1)
        self.vGutter.set(1)
        
        #define widgets
        self.removeGuidesChk = tk.Checkbutton(self, text=" Remove existing guides", variable=self.removeGuides).grid(column=0, row=1)

        self.lineHeightPt_label = tk.Label(self, text="Leading (line height) in pt: ").grid(column=0, row=2, sticky='W')
        self.xHeightPt_label = tk.Label(self, text="x-height in pt (for imagelines): ").grid(column=0, row=3, sticky='W')
        self.marginTop_label = tk.Label(self, text="Lines in top margin: ").grid(column=0, row=4, sticky='W')
        self.marginBottom_label = tk.Label(self, text="Lines in bottom margin: ").grid(column=0, row=5, sticky='W')
        self.marginLeft_label = tk.Label(self, text="Lines in left margin: ").grid(column=0, row=6, sticky='W')
        self.marginRight_label = tk.Label(self, text="Lines in right margin: ").grid(column=0, row=7, sticky='W')
        self.rowCount_label = tk.Label(self, text="Number of rows: ").grid(column=0, row=8, sticky='W')
        self.columnCount_label = tk.Label(self, text="Number of columns:").grid(column=0, row=9, sticky='W')
        self.hGutter_label = tk.Label(self, text="Lines in row gutter:").grid(column=0, row=10, sticky='W')
        self.vGutter_label = tk.Label(self, text="Lines in column gutter:").grid(column=0, row=11, sticky='W')

        self.lineHeightPt_entry = tk.Entry(self, width=12, textvariable=self.lineHeightPt).grid(column=1, row=2, padx=5)        
        self.xHeightPt_entry = tk.Entry(self, width=12, textvariable=self.xHeightPt).grid(column=1, row=3)        
        self.marginTop_entry = tk.Entry(self, width=12, textvariable=self.marginTop).grid(column=1, row=4)
        self.marginBottom_entry = tk.Entry(self, width=12, textvariable=self.marginBottom).grid(column=1, row=5)
        self.marginLeft_entry = tk.Entry(self, width=12, textvariable=self.marginLeft).grid(column=1, row=6)
        self.marginRight_entry = tk.Entry(self, width=12, textvariable=self.marginRight).grid(column=1, row=7)
        self.rowCount_entry = tk.Entry(self, width=12, textvariable=self.rowCount).grid(column=1, row=8)
        self.columnCount_entry = tk.Entry(self, width=12, textvariable=self.columnCount).grid(column=1, row=9)
        self.hGutter_entry = tk.Entry(self, width=12, textvariable=self.hGutter).grid(column=1, row=10)
        self.vGutter_entry = tk.Entry(self, width=12, textvariable=self.vGutter).grid(column=1, row=11)
        
        self.quitButton = tk.Button(self, text="Quit", command=self.quit).grid(column=0,row=12,padx=5,pady=5)
        self.calculateButton = tk.Button(self, text="Calculate", width=9, command=self.calculateButton_pressed).grid(column=1,row=12)

    def calculateButton_pressed(self):
        calculateGrid(self.removeGuides.get(),self.lineHeightPt.get(),self.xHeightPt.get(),self.marginTop.get(),self.marginBottom.get(),self.marginLeft.get(),self.marginRight.get(),self.rowCount.get(),self.columnCount.get(),self.hGutter.get(),self.vGutter.get())

    def quit(self):
        self.master.destroy()

##################################################
# Start program

def main():
    if scribus.haveDoc() == 0:
        scribus.messageBox("Error: No document open","Please, create (or open) a document before running this script ...",scribus.ICON_WARNING,scribus.BUTTON_OK)
        return
    
    try:
        scribus.statusMessage('Running script...')
        scribus.progressReset()
        unit = scribus.getUnit()
        scribus.setUnit(UNIT_POINTS)
        root = tk.Tk()
        app = TkGrid(root)
        root.mainloop()
    finally:
        if scribus.haveDoc():
            scribus.redrawAll()
        scribus.setUnit(unit)
        scribus.statusMessage('Done.')
        scribus.progressReset()

if __name__ == '__main__':
    main()
