# -*- coding: utf-8 -*-
# @Author: WuG03
# @Date:   2017-09-21 14:52:27
# @Last Modified by:   GalenMoo
# @PLast Modified time: 2017-10-24 09:56:45
import csv
import time
import mbox
import visa
import datetime
import os

from tkinter import filedialog
from tkinter import ttk
from tkinter import *

# Globals
rm = None            
resources = None     
instrument = None    

class Application(Frame):
    # The GUI and functions
    def __init__(self, master):
        Frame.__init__(self, master)
        self.root = master
        self.config()
        self.create_variables()
        self.create_widgets()
        self.grid_widgets()
        self.listOfFrames = [self.format_frame, self.init_frame, self.loadState_frame]
        self.init_grid()

        self.listOfButtons = [self.connect_btn, self.cal_btn, self.load_btn, self.sn_btn, self.collect_btn, self.file_btn, self.continue_btn]
        self.listOfLabels = [self.titleLabel, self.format_label, self.frameLabel, self.loadState_label]
        self.listOfRadioButtons = [self.RjX_rbtn, self.GjB_rbtn] 
        self.bind('<Configure>', self.resize)


    # General configuration of GUI
    def config(self):
        root.title('Feedthru Capacitor Testing')
        self.titleLabel = ttk.Label(self, text = 'Feedthru Capacitor Testing', style = "10.TLabel")
        ttk.Separator(self, orient = 'horizontal').grid(column = 0, row = 1, columnspan = 15, sticky = N+S+E+W)

        # Creates the popup window object
        self.Mbox = mbox.Mbox
        self.Mbox.root = self
        self.root.protocol("WM_DELETE_WINDOW", self.closingProgram)


    # Create variables used for logic
    def create_variables(self):
        self.vna = None
        self.VNAModel = None
        self.isVNAConnected = False
        self.isImpedanceTesting = False

        self.serialString = StringVar(self, "SN: 0")
        self.serialNum = 0

        self.file = StringVar(self, "Set Folder")
        self.isDataSelected = False
        self.currentDataRow = 4

        self.continueDisplay = StringVar(self,"Continue\nInit Data")
        self.stateNum = 0

        self.dropValue = StringVar(self)
        self.radiobtnValue = IntVar()
        self.impedanceCheckVal = IntVar(value = False)
        self.screenCapCheckVal = IntVar(value = False)

    # Creates the widgets
    def create_widgets(self):
        #----------------------------------------------- Frames -----------------------------------------------#
        self.format_label = ttk.Label(text = "Format:", style = "10.TLabel")
        self.format_frame = ttk.LabelFrame(self, labelwidget = self.format_label)

        self.frameLabel = ttk.Label(text = "Init:", style = "10.TLabel")
        self.init_frame = ttk.LabelFrame(self, labelwidget = self.frameLabel)

        self.loadState_label = ttk.Label(text = "States:", style = "10.TLabel")
        self.loadState_frame = ttk.LabelFrame(self, labelwidget = self.loadState_label)
        #----------------------------------------------- Buttons -----------------------------------------------#
        self.RjX_rbtn = ttk.Radiobutton(self.format_frame, text = "R + jX", variable = self.radiobtnValue, value = 1, style = "10.TRadiobutton")
        self.GjB_rbtn = ttk.Radiobutton(self.format_frame, text = "G + jB", variable = self.radiobtnValue, value = 2, style = "10.TRadiobutton")
        self.connect_btn = ttk.Button(self.init_frame, text = 'Connect', command = self.connectDevice, style = "10.TButton")
        self.cal_btn = ttk.Button(self.init_frame, text = 'Calibrate', command = self.initCalibrate, style = "10.TButton")
        self.load_btn = ttk.Button(self.loadState_frame, text = 'Load State', command = self.loadState, style = "10.TButton")
        self.continue_btn = ttk.Button(self, textvariable = self.continueDisplay, command = self.continueActions, width = 13, style = "10.TButton")
        self.sn_btn = ttk.Button(self, textvariable = self.serialString, command = self.serialNumberInput, style = "10.TButton")
        self.collect_btn = ttk.Button(self, text = 'Collect Data', command = self.collectData, width = 30, style = "10.TButton")
        self.file_btn = ttk.Button(self, textvariable = self.file, command = self.selectLocation, width = 30, style = "10.TButton")
        #----------------------------------------- Dropdown Menu -----------------------------------------------#
        self.states = ['Default State', 'J02', 'J03', 'J04', 'J05', 'J06', 'J08', 'J09', 'J10', 'J11', 'J12']
        self.dropDown = ttk.OptionMenu(self.loadState_frame, self.dropValue, *self.states, command = self.updateStateNum)
        self.impedanceCheck = ttk.Checkbutton(self, text = "w/ Case", variable = self.impedanceCheckVal, style = "10.TCheckbutton")

    # Grids out the widgets
    def grid_widgets(self):
    #  r/c      0               1              2              
    #   0         "Feedthru Capacitor Testing"
    #   1   -------------------------------------------
    #   2   connect_btn     load_btn      ---------        
    #   3   cal_btn         dropDown     | ctn_btn |
    #   4   GjB_rbtn        sn_btn        ---------
    #   5   RjX_rbtn            collect_btn
    #   6                 file_btn
        self.titleLabel.grid(column = 0, row = 0, columnspan = 15, sticky = N+S)       
        self.init_frame.grid(column = 0, row = 2, rowspan = 2, sticky='nsew')
        self.connect_btn.grid(column = 0, row = 0, padx = 8, pady = 8, sticky = N+S+E+W)
        self.cal_btn.grid(column = 0, row = 1, padx = 8, pady = 8, sticky = N+S+E+W)
        
        self.format_frame.grid(row = 4, column = 0, rowspan = 2, sticky = "nsew")
        self.RjX_rbtn.grid(row = 0, column = 0, padx = 8, pady = 8, sticky = N+S)
        self.GjB_rbtn.grid(row = 1, column = 0, padx = 8, pady = 8, sticky = N+S)

        self.loadState_frame.grid(column = 1, row = 2, rowspan = 2, sticky = N+S+E+W)
        self.load_btn.grid(column = 0, row = 0, padx = 8, pady = 8, sticky = N+S+E+W)
        self.dropDown.grid(column = 0, row = 1, padx = 8, pady = 8, sticky = N+S+E+W)

        self.sn_btn.grid(column = 1, row = 4, sticky = N+S+E+W)
        self.collect_btn.grid(column = 1, row = 5, columnspan = 30, sticky = N+S+E+W)
        self.file_btn.grid(column = 1, row = 6, columnspan = 30, sticky = N+S+E+W)

        self.continue_btn.grid(column = 2, row = 2, rowspan = 3, columnspan = 6, sticky = N+S+E+W)
        self.impedanceCheck.grid(row = 6, column = 0, sticky = N+S+E+W)


    # Sets the widget grids
    def init_grid(self):
        self.grid(column = 0, row = 0, sticky = 'nsew')

        for i in range(3):
            self.columnconfigure(i, weight = 1)
        for i in range(7):
            self.rowconfigure(i, weight = 1)
        
        for frame in self.listOfFrames:
            frame.columnconfigure(0, weight = 1)
            frame.rowconfigure(0, weight = 1)
            frame.rowconfigure(1, weight = 1)

        Grid.rowconfigure(root, 0, weight = 1)
        Grid.columnconfigure(root, 0, weight = 1)
        
        # Adds nice padding
        for child in self.winfo_children(): 
           child.grid_configure(padx = 8, pady = 8)
        self.sn_btn.grid(pady = 15)


    # Subfunction for destroying top popup for keybindings
    def destroyTop(self, event):
        self.top.destroy()


    # Method bound to the <Configure> event for resizing
    def resize(self, event):
        # Get geometry info from the root window.
        wm, hm = self._root().winfo_width(), self._root().winfo_height()
        
        # Random formula I guessed & checked
        fontHeight = (hm - 250) // 15
        print('Resizing to font ' + str(fontHeight))

        #--- Calculate the best font to use (use int rounding) ---#
        # Default to min size
        bestStyle_btn = styleDictionary["10.styleName_Button"]    
        bestStyle_label = styleDictionary["10.styleName_Label"]    
        bestStyle_rdbtn = styleDictionary["10.styleName_Radiobutton"]    
        bestStyle_dropmenu = styleDictionary["10.styleName_Menubutton"]    
        bestStyle_checkbtn = styleDictionary["10.styleName_Checkbutton"]

        if fontHeight < 10: pass # Min size
        elif fontHeight >= 50: # Max size
            bestStyle_btn = styleDictionary["50.styleName_Button"]    
            bestStyle_label = styleDictionary["50.styleName_Label"]    
            bestStyle_rdbtn = styleDictionary["50.styleName_Radiobutton"]    
            bestStyle_dropmenu = styleDictionary["50.styleName_Menubutton"]   
            bestStyle_checkbtn = styleDictionary["50.styleName_Checkbutton"]

        else: # Everything in between
            bestFitFont = (fontHeight // 5) * 5
            bestStyle_btn = styleDictionary[str(bestFitFont) + ".styleName_Button"]    
            bestStyle_label = styleDictionary[str(bestFitFont) + ".styleName_Label"]    
            bestStyle_rdbtn = styleDictionary[str(bestFitFont) + ".styleName_Radiobutton"]    
            bestStyle_dropmenu = styleDictionary[str(bestFitFont) + ".styleName_Menubutton"]    
            bestStyle_checkbtn = styleDictionary[str(bestFitFont) + ".styleName_Checkbutton"]

        # Set the style to the widget
        for btn in self.listOfButtons:
            btn.config(style = bestStyle_btn)
        
        for rBtn in self.listOfRadioButtons:
            rBtn.config(style = bestStyle_rdbtn)

        for label in self.listOfLabels:
            label.config(style = bestStyle_label)
        self.dropDown.config(style = bestStyle_dropmenu)
        self.impedanceCheck.config(style = bestStyle_checkbtn)

    # Connect to VNA
    def connectDevice(self):
        print("Connecting VNA")
        try:
            (self.isVNAConnected, vna_address) = self.findDevice()
            if (self.isVNAConnected):
                self.Mbox("VNA connected")
                self.vna = rm.open_resource(vna_address)
                # self.vna.write('*RST')
                self.vna.write("*CLS")
            else:
                self.Mbox("VNA cannot be connected")
        except:
            self.Mbox("VNA cannot be connected")

    # Subunction for connecting to E5062A VNA
    def findDevice(self):
        global instrument, rm, resources
        rm = visa.ResourceManager()
        resources = rm.list_resources()
        
        # Checks each connected port
        for item in resources:
            print(item)
            if "ASRL" in item:
                continue

            instrument = rm.open_resource(item)
            instrument.write_termination = '\n'
            instrument.read_termination = '\n'
            idn = instrument.query_ascii_values('*IDN?', converter = 's')
            print(idn)
            
            # If E5062A VNA, then...
            if ("E5061B" or "E5062A") in idn[1]:
                vna_address = item.encode('ascii','ignore')
                print("Found " + idn[1])
                self.VNAModel = idn[1] 
                return (True, vna_address)
                # Tuple that returns (isVNAConnected, vna_address)
            else:
                vna_address = ''
                print("Device not found")

            # Case where it can't connect
            instrument.close()
        return (False, None)  


    # Loads self.states chosen from dropdown
    def loadState(self):
        print("Loading State")
        if (self.isVNAConnected):
            stateFile = self.states[self.stateNum]
            
            self.vna.write("*CLS")
            self.vna.write(":MMEM:LOAD \"D:\\Nodestates\\%s.STA\"" %stateFile)
            errorQuery = self.vna.query_ascii_values(":SYST:ERR?", converter = 's')

            if (errorQuery[0] != "+0"):
                self.Mbox("Error " + errorQuery[1])
                return True
        else:
            self.Mbox("VNA not connected")


    # Change device serial number
    def serialNumberInput(self):
        self.top = Toplevel(self)
        self.top.grab_set() 

        ttk.Label(self.top, text = "Input Serial Number").pack(side = "top", fill = "x")
        entryText = StringVar(self, str(self.serialNum))
        self.top.snEntry = ttk.Entry(self.top, textvariable = entryText)
        self.top.snEntry.pack(side = "top", fill = "x")
        ttk.Button(self.top, text = "OK", command = self.top.destroy).pack()
        self.top.snEntry.focus_set()
        self.top.bind('<Return>', self.destroyTop)
        self.top.wait_window()

        self.serialNum = entryText.get()
        self.serialString.set("SN: " + entryText.get())


    # Collects and appends to CSV file
    def collectData(self):        
        if (self.isVNAConnected):
            if (self.isDataSelected):
                if (self.radiobtnValue.get() == 0):
                    self.Mbox("Please select a format type")
                    return

                elif (self.radiobtnValue.get() == 1):
                    self.vna.write(":CALC1:FORM SMIT")
                elif (self.radiobtnValue.get() == 2):
                    self.vna.write(":CALC1:FORM SADMittance")

                self.vna.write(":SENS1:AVER OFF")
                time.sleep(0.5)
                self.vna.write(":SENS1:AVER ON")

                Mark1Data = self.vna.query_ascii_values(":CALC1:MARK1:Y?", converter = 's')
                Mark1Data[-1] = Mark1Data[-1].strip()

                Mark2Data = self.vna.query_ascii_values(":CALC1:MARK2:Y?", converter = 's')
                Mark2Data[-1] = Mark2Data[-1].strip()
                
                Mark3Data = self.vna.query_ascii_values(":CALC1:MARK3:Y?", converter = 's')
                Mark3Data[-1] = Mark3Data[-1].strip()
                
                # try:
                SN = self.serialNum
                if (self.isImpedanceTesting):
                    currentState = self.dataTop.currentNode.get()
                    self.stateNum += 1
                    if (self.stateNum == 11):
                        self.stateNum = 1
                    self.dataTop.currentNode.set(self.dataTop.listOfNodes[self.stateNum])
                else:
                    currentState = self.states[self.stateNum]

                print(currentState)
                file = open(self.file.get(), 'a', newline = "")
                writer = csv.writer(file)
                writer.writerow((SN, currentState, Mark1Data[0], Mark1Data[1], "Calculated", "Calculated", Mark2Data[0], Mark2Data[1], 
                                    "Calculated", "Calculated", Mark3Data[0], Mark3Data[1], "Calculated", "Calculated"))
                self.updateDataWindow(currentState = currentState, data = [Mark1Data, Mark2Data, Mark3Data])
                
                if (self.screenCapCheckVal.get()):
                    currentTime = datetime.datetime.now()
                    self.vna.write("*CLS")
                    # self.vna.write(":MMEM:STOR:IMAG \"D:\\ScreenCapture\\" + str(currentTime)[:10] + "\\" + currentState + "_" + str(currentTime)[11:-10].replace(":", ".") + ".bmp\"")
                    s1pFileLocation = "\"D:\\ScreenCapture\\" + str(currentTime)[:10] + "_" + self.serialNum + "\\" + currentState + "_" + str(currentTime)[11:-10].replace(":", ".") + ".s1p\""
                    S1PFileName = currentState + "(" + str(currentTime)[11:-10].replace(":", ".") + ")" + ".s1p"

                    self.vna.write(":MMEM:STOR:SNP " + s1pFileLocation)

                    errorQuery = self.vna.query_ascii_values(":SYST:ERR?", converter = 's')
                    print(errorQuery)

                    if (errorQuery[0] != "+0"):
                        print("Creating screen capture directory")
                        self.vna.write(":MMEM:MDIR \"D:\\ScreenCapture\"")
                        self.vna.write(":MMEM:MDIR \"D:\\ScreenCapture\\" + str(currentTime)[:10] + "_" + self.serialNum + "\"")
                        self.vna.write(":MMEM:STOR:SNP " + s1pFileLocation)
                        # self.vna.write(":MMEM:STOR:IMAG \"D:\\ScreenCapture\\" + str(currentTime)[:10] + "\\" + currentState + "_" + str(currentTime)[11:-10].replace(":", ".") + ".bmp\"")
                        self.vna.write("*CLS")

                    time.sleep(0.5)
                    localS1PLocation = "C:\\Users\\Public\\s1p\\" + self.serialNum + "[" + str(currentTime)[:10] + "]\\"+ S1PFileName                
                    os.makedirs(os.path.dirname(localS1PLocation), exist_ok=True)

                    S1PData = self.vna.query_ascii_values(":MMEM:TRAN? " + s1pFileLocation, converter = 's')                    
                    s1pFile = open(localS1PLocation, "w")

                    for cell in S1PData:
                        s1pFile.write(cell)
                    s1pFile.close()
                    
                    self.Mbox("S1P saved in \"" + localS1PLocation + "\"")
                else:
                    self.Mbox("Wrote data")
            else:
                self.Mbox("Select a folder")
        else:
            self.Mbox("VNA not connected")

    # Updates the data window
    def updateDataWindow(self, currentState = None, data = None):
        # data = [[1,2],[3,4],[5,6]]
        ttk.Label(self.dataTop, text = self.serialNum).grid(row = self.currentDataRow, column = 0)
        ttk.Label(self.dataTop, text = currentState).grid(row = self.currentDataRow, column = 1)
        ttk.Label(self.dataTop, text = data[0][0]).grid(row = self.currentDataRow, column = 3)
        ttk.Label(self.dataTop, text = data[0][1]).grid(row = self.currentDataRow, column = 4)
        ttk.Label(self.dataTop, text = data[1][0]).grid(row = self.currentDataRow, column = 5)
        ttk.Label(self.dataTop, text = data[1][1]).grid(row = self.currentDataRow, column = 6)
        ttk.Label(self.dataTop, text = data[2][0]).grid(row = self.currentDataRow, column = 7)
        ttk.Label(self.dataTop, text = data[2][1]).grid(row = self.currentDataRow, column = 8)
        spacing = 300 + 15*(self.currentDataRow-4)

        self.dataTop.geometry('1000x' + str(spacing))       
        self.currentDataRow += 1


    # Continue button to allow easy data intake. Process is as follows:
    #   - Load next state (as seen on the button)
    #   - Collect data
    #   - Update labels
    def continueActions(self):
        if (self.isVNAConnected):
            if (self.isDataSelected):
                if (self.continueDisplay.get() == "Continue\nInit Data"):
                    self.Mbox("Please select a file")
                    return

                if (self.radiobtnValue.get() == 0):
                    self.Mbox("Please select a format type")
                    return

                if (self.stateNum == 10):
                    self.serialNumberInput()

                self.stateNum += 1
                if (self.stateNum == 11):
                    self.stateNum = 1
                self.continueDisplay.set("Continue\nState: " + self.states[self.stateNum])
                # print(self.stateNum)    
                
                isError = self.loadState()
                if (isError == True):
                    return
                
                if (self.stateNum == 10):
                    self.continueDisplay.set("Continue\nState: " + self.states[1])
                else:
                    self.continueDisplay.set("Continue\nState: " + self.states[self.stateNum + 1])
                time.sleep(2)
                self.collectData()
            else:
                self.Mbox("Select a folder")
        else:
            self.Mbox("VNA not connected")


    # Pick a file location or existing file
    def selectLocation(self):
        if (self.isVNAConnected):
            if (self.radiobtnValue.get() == 1):
                unit = "ohm"
            elif (self.radiobtnValue.get() == 2):
                unit = "S"
            else:
                self.Mbox("Select a format type")
                return
            if (self.impedanceCheckVal.get() == True):
                self.isImpedanceTesting = True
                self.stateNum = 1
            else:
                self.isImpedanceTesting = False
                
            isNewFile = self.getFileMethod()
            if (isNewFile):
                folderName = filedialog.askdirectory()
                print(folderName)
                if (folderName == ""):
                    return

                time = datetime.datetime.now()

                if (self.impedanceCheckVal.get() == True):
                    fileName = ("/ImpedanceMeasurements[" + str(time)[:10] + "][" + str(time)[11:-10] + "].csv").replace(":", ".")

                else:
                    fileName = ("/FeedthruTesting[" + str(time)[:10] + "][" + str(time)[11:-10] + "].csv").replace(":", ".")
                print(fileName)

                with open(folderName + fileName, "w", newline = '') as file:
                    writer = csv.writer(file)
                    writer.writerow(("SN", "Node", "1 MHz", "", "", "", "64 MHz", "", "", "", "128 MHz", "", "", "", "Data collected on " + str(time)[:-10]))
                    writer.writerow((""  , ""    , "Real ({})".format(unit), "Imaginary ({})".format(unit), "Cap (nF)", "Z (ohm)", "Real ({})".format(unit), 
                                        "Imaginary ({})".format(unit), "Cap (nF)", "Z (ohm)", "Real ({})".format(unit), "Imaginary ({})".format(unit), "Cap (nF)", "Z (ohm)"))                
                    
                self.file.set(folderName + fileName)            
                self.setupDataWindow()

            else:   # Existing file
                fileName = filedialog.askopenfilename()
                if (fileName == ""):
                    return

                print(fileName)
                with open(fileName, "r", newline = '') as file:
                    reader = csv.reader(file)
                    next(reader, None)
                    x = next(reader)[2]
                    if (x[6] == 'o'):
                        self.radiobtnValue.set(1)                    
                    else:
                        self.radiobtnValue.set(2)

                    self.setupDataWindow()
                    rowCount = 3
                    for row in reader:
                        self.serialNum = row[0]
                        self.serialString.set("SN: " + row[0])
                        try:
                            self.stateNum = self.states.index(row[1])
                        except:
                            print("Can not find Node of row " + str(rowCount))
                        readData = [[row[2], row[3]], [row[6], row[7]], [row[10], row[11]]]
                        self.updateDataWindow(data = readData)
                        rowCount += 1
                        print(row)
                self.file.set(fileName)
            if (self.stateNum == 10):
                self.continueDisplay.set("Continue\nState: " + self.states[1])
            else:            
                self.continueDisplay.set("Continue\nState: " + self.states[self.stateNum+1])
            self.isDataSelected = True
        else:
            self.Mbox("VNA not connected")

    # Subfunction to determine file method
    def getFileMethod(self):
        fileOption = IntVar()

        self.top = Toplevel(self)
        ttk.Label(self.top, text = "Select File Method").grid(column = 0, row = 0)
        ttk.Separator(self, orient = 'horizontal').grid(column = 0, row = 1, columnspan = 15, sticky = 'ew')

        ttk.Radiobutton(self.top, text = "Existing File", variable = fileOption, value = False).grid(column = 0, row = 2, padx = 5, pady = 5)
        ttk.Radiobutton(self.top, text = "New File", variable = fileOption, value = True).grid(column = 1, row = 2, padx = 5, pady = 5)
        enter_btn = ttk.Button(self.top, text = "Enter", command = self.top.destroy, width = 30).grid(column = 0, row = 3, columnspan = 30)
        
        self.top.bind('<Return>', self.destroyTop)
        self.top.wait_window()
        return fileOption.get()

    # Inits the data window
    def setupDataWindow(self):
        if (self.radiobtnValue.get() == 1):
            unit = "ohm"
        elif (self.radiobtnValue.get() == 2):
            unit = "S"  
        else:
            unit = "unitless"

        self.dataTop = Toplevel(self)
        self.dataTop.geometry('1000x300')       
        
        for i in range(10):
            self.dataTop.columnconfigure(i, weight = 1)

        ttk.Label(self.dataTop, text = "Data Set").grid(column = 0, row = 0, columnspan = 15, sticky = N+S)     
        ttk.Label(self.dataTop, text = "1MHz").grid(column = 3, row = 1, columnspan = 2, sticky = N+S) 
        ttk.Label(self.dataTop, text = "64MHz").grid(column = 5, row = 1, columnspan = 2, sticky = N+S) 
        ttk.Label(self.dataTop, text = "128MHz").grid(column = 7, row = 1, columnspan = 2, sticky = N+S) 

        ttk.Label(self.dataTop, text = "SN").grid(column = 0, row = 2)
        ttk.Label(self.dataTop, text = "Node").grid(column = 1, row = 2)
        ttk.Label(self.dataTop, text = "Real ({})".format(unit)).grid(column = 3, row = 2)
        ttk.Label(self.dataTop, text = "Imaginary ({})".format(unit)).grid(column = 4, row = 2)
        ttk.Label(self.dataTop, text = "Real ({})".format(unit)).grid(column = 5, row = 2)
        ttk.Label(self.dataTop, text = "Imaginary ({})".format(unit)).grid(column = 6, row = 2)
        ttk.Label(self.dataTop, text = "Real ({})".format(unit)).grid(column = 7, row = 2)
        ttk.Label(self.dataTop, text = "Imaginary ({})".format(unit)).grid(column = 8, row = 2)
        ttk.Label(self.dataTop, text = "Delete").grid(column = 9, row = 2)

        ttk.Separator(self.dataTop, orient = "horizontal").grid(row = 3, columnspan = 15, sticky = N+S+E+W)
        ttk.Separator(self.dataTop, orient = "vertical").grid(row = 4, column = 2, rowspan = 500, sticky = N+S)
        # ttk.Button(self.dataTop, text = "Press me", command = self.updateDataWindow).grid(row = 0, column = 0)
        ttk.Checkbutton(self.dataTop, text = "Screen Capture", variable = self.screenCapCheckVal).grid(row = 0, column = 8, sticky = N+S)
        if (self.isImpedanceTesting):
            self.dataTop.currentNode = StringVar(self)
            self.dataTop.listOfNodes = ['Default', 'LVtip', 'RVtip', 'LVR1', 'RVring', 'Atip', 'RVC', 'LVR2', 'SVC', 'LVR3', 'Aring']
            self.dataTop.nodeMenu = ttk.OptionMenu(self.dataTop, self.dataTop.currentNode, *self.dataTop.listOfNodes, command = self.updateStateNum)
            self.dataTop.nodeMenu.grid(row = 0, column = 0, columnspan = 2, rowspan = 2, sticky = N+S+E+W)
            self.dataTop.currentNode.set('LVtip')

    # Create calibration menu, ECal or 1 Port Cal
    def initCalibrate(self):
        if (self.isVNAConnected):
            self.top = Toplevel(self)
            self.top.grab_set()
            ttk.Button(self.top, text = "E-Cal", command = self.eCalibration).pack(fill = 'both', expand = True)
            ttk.Button(self.top, text = "1 Port Calibration", command = self.onePortCalibration).pack(fill = 'both', expand = True)
        else:
            self.Mbox("VNA not connected")

    # ECal with 58 sec progressbar
    def eCalibration(self):
        self.setMeasureSettings()
        self.top.destroy()
        self.top = Toplevel(self)
        self.top.grab_set()

        if (self.VNAModel == "E5062A"):
            self.timeDuration = 58000
        else:
            self.timeDuration = 3000

        self.top.progressBar = ttk.Progressbar(self.top, orient = "horizontal", length = 200, mode = "determinate", value = 0, maximum = self.timeDuration)
        self.top.progressBar.pack()
        self.vna.write(":SENS1:CORR:COLL:ECAL:SOLT1 1")
        self.top.currentProgress = 0
        self.updateProgress()

    # Update progressbar every 100ms
    def updateProgress(self):
        self.top.currentProgress += 100
        self.top.progressBar['value'] = self.top.currentProgress
        if self.top.currentProgress < self.timeDuration:
            self.top.after(100, self.updateProgress)
        else:
            self.top.destroy()
            self.Mbox("ECal complete")


    def onePortCalibration(self):
        def openCal():
            self.vna.write(":SENS1:CORR:COLL:OPEN 1")
            self.calCounter[0][1] = True
        
        def shortCal():
            self.vna.write(":SENS1:CORR:COLL:SHOR 1")
            self.calCounter[1][1] = True

        def loadCal():
            self.vna.write(":SENS1:CORR:COLL:LOAD 1")
            self.calCounter[2][1] = True

        def saveCal():
            errorString = "Incomplete calibration of "
            firstCalFlag = True
            for calType in range(3):
                if self.calCounter[calType][1] == True:
                    continue
                if (firstCalFlag):
                    errorString += self.calCounter[calType][0]
                    firstCalFlag = False
                else:
                    errorString += ", " + self.calCounter[calType][0]
            if (not(firstCalFlag)):
                self.Mbox(errorString + ".")
                return
            self.top.destroy()
            self.top = Toplevel(self)
            self.top.grab_set() 
            entryText = StringVar()

            ttk.Label(self.top, text = "Input file name").pack(side = "top", fill = "x")
            self.top.snEntry = ttk.Entry(self.top, textvariable = entryText)
            self.top.snEntry.pack(side = "top", fill = "x")
            ttk.Button(self.top, text = "OK", command = self.top.destroy).pack()
            self.top.bind('<Return>', self.destroyTop)
            self.top.snEntry.focus_set()
            self.top.wait_window()
            fileName = entryText.get()

            self.vna.write("*CLS")
            self.vna.write(":MMEM:STOR \"D:\\Nodestates\\%s.STA\"" %fileName)
            errorQuery = self.vna.query_ascii_values(":SYST:ERR?", converter = 's')
            print(errorQuery)

            if (errorQuery[0] != "+101"):
                print("Creating Nodestates directory")
                self.vna.write(":MMEM:MDIR \"D:\\NodeStates\"")
                self.vna.write(":MMEM:STOR \"D:\\Nodestates\\%s.STA\"" %fileName)
                self.vna.write("*CLS")

            self.Mbox("State saved in D:\\Nodestates\\%s.STA" %fileName)

        self.top.destroy()
        self.setMeasureSettings()
        self.calTop = Toplevel(self)
        self.calTop.grab_set()
        self.calCounter = [["open", False], ["short", False], ["load", False]]

        ttk.Button(self.calTop, text = "Open", command = openCal).pack(side = "top", fill = "x")
        ttk.Button(self.calTop, text = "Short", command = shortCal).pack(side = "top", fill = "x")
        ttk.Button(self.calTop, text = "Load", command = loadCal).pack(side = "top", fill = "x")
        ttk.Button(self.calTop, text = "Save", command = saveCal).pack(side = "top", fill = "x")

    # Setting markers and cal settings
    def setMeasureSettings(self):
        # start, stop, center, span and format
        self.vna.write(":SENS1:FREQ:STAR 1E6")
        self.vna.write(":SENS1:FREQ:STOP 200E6")
        self.vna.write(":SENS1:FREQ:CENT 100.5E6")
        self.vna.write(":SENS1:FREQ:SPAN 199E6")
        self.vna.write(":CALC1:FORM SMIT")

        # 3 markers at 1MHz, 64MHz, and 128MHz
        self.vna.write(":CALC1:MARK1 ON")
        self.vna.write(":CALC1:MARK1:X 1E6")
        self.vna.write(":CALC1:MARK2 ON")
        self.vna.write(":CALC1:MARK2:X 64E6")
        self.vna.write(":CALC1:MARK3 ON")
        self.vna.write(":CALC1:MARK3:X 128E6")

        self.vna.write(":INIT1:CONT ON")

    def updateStateNum(self, event):
        if (self.isImpedanceTesting):
            try:
                self.stateNum = self.dataTop.listOfNodes.index(self.dataTop.currentNode.get())
            except:
                self.Mbox("Error: data dialog not open")
        else:
            self.stateNum = self.states.index(self.dropValue.get())
            if (self.stateNum == 10):
                self.continueDisplay.set("Continue\nState: " + self.states[1])
            else:
                self.continueDisplay.set("Continue\nState: " + self.states[self.stateNum + 1])
        print(self.stateNum)
    
    # Window close event function
    def closingProgram(self):
        global instrument, rm
        try:
            if (self.isVNAConnected):
                print("Closing instrument")
                instrument.close()
        except:
            pass
        print("Closing window")
        self.root.destroy()
        self.quit

if __name__ == "__main__":
    root = Tk()             # Creates the framework
    
    # Inits various font sizes for font scaling
    fontList = ['Button', 'Label', 'Radiobutton', 'Menubutton', 'Checkbutton']
    styleDictionary = {}
    for i in range(len(fontList)):
        for font in range(10, 51, 5):
            styleDictionary[str(font) + ".styleName_" + fontList[i]] = str(font) + '.T' + fontList[i]
            fontName = ' '.join(['Helvetica', str(font)])
            ttk.Style().configure(styleDictionary[str(font) + ".styleName_" + fontList[i]], font = fontName)

    Application(root)       # GUI object
    root.mainloop()         # Mainloop from the frame