""" 
    This tool is intended to be interfaced with the Shutter Lover measurement tool.
    The value of such an application is to help the user to enter series of measurement values 
    into a spreadsheet application. 
    So that, the user will be easily able to make statistics, mean value, mean deviation and so on,
    and also make archives of the measurements results.
"""
__author__ = "Sebastien ROY"
__license__ = "GPL"
__version__ = "1.1.0"
__status__ = "Released"

import threading
import time
import queue
import json
import sys
from json import JSONDecodeError
from tkinter import *
from  tkinter import ttk
import serial.tools.list_ports
from serial import SerialException
import datetime
from dataclasses import dataclass

is_stopped = False
DEBUG = True
app = None

# we define that the speed setting is in col 1 (the 2nd column)
SPEED_SETTING_COL = 1
SPEED_SETTING_ID = 'setting'

def testMeasureThread():
    """ Test purpose:
        Test listener used as a mock for COM port listener 
    """
    global app, is_stopped
    index = 0
    currentId = 0
    entries = ['Tagada tsouin', 
        '{\"eventType\": \"MultiSensorMeasure\", \"unit\": \"microsecond\", \"firmware_version\": \"1.0.0\",\
            \"bottomLeftOpen\": 0, \"bottomLeftClose\": 987, \"centerOpen\": 3456, \"centerClose\": 4567, \"topRightOpen\": 5678, \"topRightClose\": 6789,\
            \"bottomLeftOpenOffset\":-25, \"bottomLeftCloseOffset\":32, \"topRightOpenOffset\":40, \"topRightCloseOffset\":10 }',
        '{\"eventType\": \"MultiSensorMeasure\", \"unit\": \"microsecond\", \"firmware_version\": \"1.0.0\",\
            \"bottomLeftOpen\": 0, \"bottomLeftClose\": 987, \"centerOpen\": 3456, \"centerClose\": 4567, \"topRightOpen\": -1, \"topRightClose\": -1,\
            \"bottomLeftOpenOffset\":-25, \"bottomLeftCloseOffset\":32, \"topRightOpenOffset\":40, \"topRightCloseOffset\":10 }',
        '{\"eventType\": \"MultiSensorMeasure\", \"unit\": \"microsecond\", \"firmware_version\": \"1.0.0\",\
            \"bottomLeftOpen\": -1, \"bottomLeftClose\": -1, \"centerOpen\": 0, \"centerClose\": 987, \"topRightOpen\": 3456, \"topRightClose\": 4567,\
            \"bottomLeftOpenOffset\":-25, \"bottomLeftCloseOffset\":32, \"topRightOpenOffset\":40, \"topRightCloseOffset\":10 }',
        '{\"eventType\": \"MultiSensorMeasure\", \"unit\": \"microsecond\", \"firmware_version\": \"1.0.0\",\
            \"bottomLeftOpen\": 0, \"bottomLeftClose\": 987, \"centerOpen\": -1, \"centerClose\": -1, \"topRightOpen\": 5678, \"topRightClose\": 6789,\
            \"bottomLeftOpenOffset\":-25, \"bottomLeftCloseOffset\":32, \"topRightOpenOffset\":40, \"topRightCloseOffset\":10 }',
        '{\"eventType\": \"MultiSensorMeasure\", \"unit\": \"microsecond\", \"firmware_version\": \"1.0.0\",\
            \"bottomLeftOpen\": 0, \"bottomLeftClose\": 987, \"centerOpen\": 3456, \"centerClose\": 4567, \"topRightOpen\": 5678, \"topRightClose\": -1,\
            \"bottomLeftOpenOffset\":-25, \"bottomLeftCloseOffset\":32, \"topRightOpenOffset\":40, \"topRightCloseOffset\":10 }',
        'pof pof', 
         ''] 
    while not is_stopped:
        try:
            data = json.loads(entries[index])
            # Put a data in the queue
            app.comque.put(data)
            # Generate an event in order to notify the GUI
            app.ws.event_generate('<<Measure>>', when='tail')       
        except JSONDecodeError:
            if(DEBUG):
                print("Not JSon : " + entries[index])
        except TclError:
            break  
        time.sleep(2)
        index += 1
        if (index>= len(entries)):
            index = 0

def measureThread():
    """ This function is executed in a separated thread 
        When JSon data is read on the serial port,
        the function pushs a <<Measure>> event in a queue for the main thread
    """
    global app, is_stopped
    try:
        while not is_stopped:
            if(app.serialPort == None or app.serialPort.is_open == False):
                time.sleep(1)
                app.openSerialPort(app.portName)
                if(app.serialPort == None):
                    app.connectionStatusLabel.set("No connection")
                elif(app.serialPort.is_open): 
                    app.connectionStatusLabel.set("Connected") 
                else:
                    app.connectionStatusLabel.set("Disconnected")

            if(app.serialPort != None and app.serialPort.is_open):
                try:
                    line = app.serialPort.readline()
                    if(DEBUG):
                        print(line)
                    data = json.loads(line)
                    # Put a data in the queue
                    app.comque.put(data)
                    # Generate an event in order to notify the GUI
                    print("Processed data event : " + str(data))
                    app.ws.event_generate('<<Measure>>', when='tail')       
                except JSONDecodeError:
                    if(DEBUG):
                        print("Not formated data:" + str(line))
                except SerialException:
                    # Connection failed
                    app.connectionStatusLabel.set("Not connected")
                    app.serialPort = None
                    app.serialPorts.pop(app.portName)
                except TclError:
                    print("TclError occured")
                except TypeError:
                    print("TypeError occured")
                except Exception:
                    print("Other error occured")
    finally:
        if(DEBUG):
            print("Measurement worker ended")

@dataclass
class DataDef:
    """
        This class is used to manage the display of the data.
        Each instance of this class describes a table column, and how to compute the displayed data
        The idea is to use raw data got from the json stream and to add calculated data into the json dictionnary
    """
    # id of the display column and key of the added information
    id: str
    # label of the column
    colLabel: str
    # Display width of the column
    colWidth: int
    # Formating string for the data. Example : "{:0.1f}" to display a float with 1 decimal
    format: str
    # the function used to compute the column data. It can use as parameter the data contained into the json dictionary
    formula: callable

    def strValue(self, value):
        if(type(value) == float):
            return self.format.format(value)
        else:
            return str(value)

def microSpeed(val1, val2, offset1=0, offset2=0):
    '''
        Computes speed in 1/s from microseconds 
        a data < 0 mean no value
    '''
    print("Offset1 = {}, Offset2 = {}\n".format(offset1, offset2))
    if(((val1+offset1) == (val2+offset2)) or (val1 < 0) or (val2 < 0)):
        return "-"
    else:
        return abs(1000000.0/((val1+offset1)-(val2+offset2)))

def microTime(val1, val2, offset1=0, offset2=0):
    '''
        Computes the time difference in ms from microseconds
        a data < 0 mean no value
    '''   
    if((val1 < 0) or (val2 < 0)):
        return "-"
    else:
        return abs(((val1+offset1)-(val2+offset2))/1000.0)

def extrapolate(val):
    '''
        Extrapolation of the course time on the whole 24x36 frame
    '''
    global app
    if(type(val) == float):
        return val * app.extrapolation_factor
    else:
        return val

class TreeviewEdit(ttk.Treeview):
    ''' This is a redefinition of the TreeView widget in order to make it editable 
        Thank you to JobinPy on YouTube for his tutorial
    '''
    def __init__(self, master, **kw):
        super().__init__(master, **kw)       
        self.bind("<Double-Button-1>", self.on_double_click)
        self.root = master
        
    def on_double_click(self, event):
        region_clicked = self.identify_region(event.x, event.y)
        if (region_clicked != 'cell'):
            return
        column = self.identify_column(event.x)
        column_index = int(column[1:])
        if(column_index != (SPEED_SETTING_COL + 1)):
            return
        selected_iid = self.focus()
        selected_values = self.item(selected_iid)
        selected_text = selected_values.get("values")[SPEED_SETTING_COL]
        column_box = self.bbox(selected_iid, column)
        entry_edit = ttk.Entry(self, width=column_box[2])
        # Record col index and item iid
        entry_edit.editing_column_index = column_index
        entry_edit.editing_item_iid = selected_iid
        
        entry_edit.insert(0, selected_text)
        entry_edit.select_range(0, len(str(selected_text)))
        entry_edit.focus()
        
        entry_edit.bind("<Escape>", self.on_escape)
        entry_edit.bind("<FocusOut>", self.update_cell)
        entry_edit.bind("<Return>", self.update_cell)
        
        entry_edit.place(x=column_box[0], y=column_box[1], w=column_box[2], h=column_box[3])
        
    def on_escape(self, event):
        event.widget.destroy()
         
    def update_cell(self, event):
        new_text = event.widget.get()
        selected_iid = event.widget.editing_item_iid
        column_index = event.widget.editing_column_index
        current_values = self.item(selected_iid).get("values")
        current_values[column_index-1] = new_text
        app.document[int(selected_iid)][SPEED_SETTING_ID] = new_text
        self.item(selected_iid, values=current_values)
        event.widget.destroy()
        

class RemoteApp:
    '''
        This is the application class
    '''

    def __init__(self, test=False):  
        self.test = test      
        self.ws = Tk()
        self.comque = queue.Queue()
        self.document = []
        self.measureId = 0
        self.serialPort = None
        self.serialPorts = {}
        self.portName = ''
        self.is_stopped = False
        self.selectedPort = StringVar()
        self.connectionStatusLabel = None
        self.connectionCombo = None
        self.selectedDirection = StringVar()
        self.extrapolation_factor = 24.0 / 20.0 # extrapolation from 20mm to 24mm (vertical direction)
        self.speedSetting = StringVar(value='60')

        # This is the columns definition : column and data id (internal identification of the culmn, must be unique), column name, column width, value format, value computation function that use json data
        # You can freely change the order, or even remove you the column of your choice.
        # To remove a column, you can either remove the related line or change it as a comment, by adding a '#' character at the beginning of the line
        self.dataDefs = [
            DataDef('id', 'Id', 20, '{}', lambda data: data['id']),
            DataDef(SPEED_SETTING_ID,'Setting (1/s)', 70, '{:0.1f}', lambda data: app.speedSetting.get()),
            DataDef('speed_c', 'Speed (1/s)', 60, '{:0.1f}', lambda data: microSpeed(data['centerClose'], data['centerOpen'])),
            DataDef('time_c', 'Time (ms)', 60, '{:0.3f}', lambda data: microTime(data['centerClose'], data['centerOpen'])),
            DataDef('course_1', 'Open (ms)', 60, '{:0.2f}', lambda data: microTime(data['topRightOpen'], data['bottomLeftOpen'], data['topRightOpenOffset'], data['bottomLeftOpenOffset'])),
            DataDef('course_2', 'Close (ms)', 60, '{:0.2f}', lambda data: microTime(data['topRightClose'], data['bottomLeftClose'], data['topRightCloseOffset'], data['bottomLeftCloseOffset'])),
            DataDef('course_1_ext', 'Open ext', 60, '{:0.2f}', lambda data: extrapolate(data['course_1'])),
            DataDef('course_2_ext', 'Close ext', 60, '{:0.2f}', lambda data: extrapolate(data['course_2'])),
            DataDef('speed_bl', 'Speed Bot. L.', 70, '{:0.1f}', lambda data: microSpeed(data['bottomLeftClose'], data['bottomLeftOpen'], data['bottomLeftCloseOffset'], data['bottomLeftOpenOffset'])),
            DataDef('time_bl', 'Time Bot. L.', 70, '{:0.3f}', lambda data: microTime(data['bottomLeftClose'], data['bottomLeftOpen'], data['bottomLeftCloseOffset'], data['bottomLeftOpenOffset'])),
            DataDef('speed_tr', 'Speed Top R.', 70, '{:0.1f}', lambda data: microSpeed(data['topRightClose'], data['topRightOpen'], data['topRightCloseOffset'], data['topRightOpenOffset'])),
            DataDef('time_tr', 'Time Top R.', 70, '{:0.3f}', lambda data: microTime(data['topRightClose'], data['topRightOpen'], data['topRightCloseOffset'], data['topRightOpenOffset'])),
            DataDef('course_1_12', 'Open 1/2', 60, '{:0.2f}', lambda data: microTime(data['centerOpen'], data['bottomLeftOpen'], 0, data['bottomLeftOpenOffset'])),
            DataDef('course_1_22', 'Open 2/2', 60, '{:0.2f}', lambda data: microTime(data['centerOpen'], data['topRightOpen'], 0, data['topRightOpenOffset'])),
            DataDef('course_2_12', 'Close 1/2', 60, '{:0.2f}', lambda data: microTime(data['centerClose'], data['bottomLeftClose'], 0, data['bottomLeftCloseOffset'])),
            DataDef('course_2_22', 'Close 2/2', 60, '{:0.2f}', lambda data: microTime(data['centerClose'], data['topRightClose'], 0, data['topRightCloseOffset']))
            ]

    def openSerialPort(self, name):
        """ Open default serial port
            If no serial port name has been defined, open the first port available
        """    
        global app

        if(self.portName == '--'):
            self.serialPort = None
        elif(name in self.serialPorts):
            self.serialPort = self.serialPorts[name]
        else:
            for info in serial.tools.list_ports.comports():
                if info.name == self.portName:
                    app.connectionStatusLabel.set("Connecting...")
                    self.serialPort = serial.Serial(info.device)
                    if(self.serialPort.is_open):
                        self.serialPorts[app.portName] = self.serialPort   
                        if(DEBUG):
                            print("SerialPort : " + str(app.serialPort))       
     
    def listSerialPorts(self):
        """ All is in the title
        """
        portNames= [info.name for info in serial.tools.list_ports.comports()]
        if not portNames:
            portNames.append("--")
        return portNames

    def handleMultiSensorMeasure(self, data):
        # for each the dataDef definition, compute its value according to the json data dictionnary
        # and add the resulting value to the dictionnary with the defined key
        for dataDef in self.dataDefs:
            data[dataDef.id] = dataDef.formula(data)
        # add the json dictionnary and its added values to the document
        self.document.append(data)
        # add a new row in the table
        tree = self.ws.nametowidget("mainFrame.measureTable")
        # iterate on all dataDef definitions, get the value that is stored with the defined id and get it as string
        rowValues = [(dataDef.strValue(data[dataDef.id])) for dataDef in self.dataDefs]
        item = tree.insert(parent='', index='end', iid=data['id'],text='', values=rowValues)
        tree.see(item)


    def dataEvent(self, event):
        """ Dispatches the events received from the listener """
        # get event data
        data = self.comque.get()
        # add an identifier
        data['id'] = self.measureId
        self.measureId += 1
        
        # dispatch according to the type
        eventType = data['eventType']
        #only one type of json event is managed
        if(eventType == 'MultiSensorMeasure'):
            self.handleMultiSensorMeasure(data)

    def clearAll(self):
        """ Callback used to clear all entries """
        # Clear treeView table
        tree = self.ws.nametowidget("mainFrame.measureTable")
        tree.delete(*tree.get_children())
        # Clear document
        self.document.clear()
        # Reset counter
        self.measureId = 0

    def string_out(self, rows, separator='\t', line_feed='\n'):
        """ Prepares a string to send to the clipboard. """
        out = []
        for row in rows:
            out.append( separator.join( row ))  # Use '\t' (tab) as column seperator.
        return line_feed.join(out)              # Use '\n' (newline) as row seperator.

    def documentToLists(self):
        """ Copy data to clipboard purpose:
            Get the document data and metadata into a list of lists (rows/columns)
        """
        data = []
        # Metadata
        data.append(["Shutter Lover data exported using Remote Application"])
        data.append(["https://github.com/sebastienroy/shutter_lover_remote_app"])
        data.append(["Date :", str(datetime.datetime.now())])
        # Columns header
        headers = [ (dataDef.colLabel) for dataDef in self.dataDefs]
        data.append(headers)
        # Data
        for row in self.document:
            # extracts all the values from the row according to the dataDef.id keys and format them using strValue function of the dataDef
            rowValues = [(dataDef.strValue(row[dataDef.id])) for dataDef in self.dataDefs]
            data.append(rowValues)
        return data
    
    def copyToClipboard(self):
        """ Copies all metadata and data to the clipboard. """
        data = self.documentToLists()
        self.ws.clipboard_clear()
        self.ws.clipboard_append(self.string_out(data, separator='\t'))     # Paste string to the clipboard
        self.ws.update()   # The string stays on the clipboard after the window is closed
        
    def on_closing(self):
        """ Close ports, exit threads... """
        global is_stopped
        is_stopped = True
        self.ws.destroy()
        if self.serialPort and self.serialPort.is_open:
            self.serialPort.close()
        sys.exit()

    def on_combo_selection(self, event):
        """ COM port selection Callback """       
        self.portName = self.selectedPort.get()
        if(DEBUG):
            print("Change port to : " + self.portName)
        self.openSerialPort(self.portName)  

    def update_cb_list(self):
        self.connectionCombo['values'] = self.listSerialPorts()

    def on_direction_selection(self, event):
        if(self.selectedDirection.get() == "Vertical"):
            self.extrapolation_factor = 24.0 / 20.0
        else:
            self.extrapolation_factor = 36.0 / 32.0
        print("Selected Direction: {}, extrapolation factor = {}".format(self.selectedDirection.get(), self.extrapolation_factor) )

    def run(self):
        """ Initialize and loop """

        self.ws.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.ws.title("Shutter Lover Remote Application")
        self.ws.geometry("1024x200")	
        frame = Frame(self.ws, name="mainFrame")
        frame.pack(expand=True, fill='both')

        #scrollbar
        v_scroll = Scrollbar(frame, name = "v_scroll")
        
        # Button_frame is necessary to put many buttons in a row
        button_frame = Frame(frame, name="buttonFrame", )

        Button(button_frame, text="Clear all", command=self.clearAll).grid(row=0, column=0, padx=5, pady=5)
        Button(button_frame, text="Copy to Clipboard", command=self.copyToClipboard).grid(row=0, column=1, padx=5, pady=5)

        # Connection
        ttk.Separator(master=button_frame, orient=VERTICAL, style='TSeparator', class_= ttk.Separator,takefocus= 0).grid(row=0, column=2, padx=5, pady=0)
        Label(button_frame, text="Device :").grid(row=0, column=3, padx=5, pady=5)
        # serial ports combo
        self.connectionCombo = ttk.Combobox(button_frame, textvariable = self.selectedPort, state='readonly', width=8,  postcommand = self.update_cb_list)
        self.ports=self.listSerialPorts()
        self.portName = self.ports[0]
        self.connectionCombo['values']=self.ports
        self.connectionCombo.current(0)
        self.connectionCombo.bind("<<ComboboxSelected>>", self.on_combo_selection)
        self.connectionCombo.grid(row=0, column=4, padx=5, pady=5)
        
        # serial port status
        self.connectionStatusLabel = StringVar(frame, "Unknown Status")
        Label(button_frame, textvariable=self.connectionStatusLabel).grid(row=0, column=5, padx=5, pady=5)

        # curtain translation direction combo
        ttk.Separator(master=button_frame, orient=VERTICAL, style='TSeparator', class_= ttk.Separator,takefocus= 0).grid(row=0, column=6, padx=5, pady=0)
        Label(button_frame, text="Curtain translation direction:").grid(row=0, column=7, padx=7, pady=5)
        directionCombo = ttk.Combobox(button_frame, values=['Vertical', 'Horizontal'], textvariable = self.selectedDirection, state='readonly', width=8,  postcommand = self.update_cb_list)
        directionCombo.current(0)
        directionCombo.bind("<<ComboboxSelected>>", self.on_direction_selection)
        directionCombo.grid(row=0, column=8, padx=5, pady=5)

        # Camera speed setting
        ttk.Separator(master=button_frame, orient=VERTICAL, style='TSeparator', class_= ttk.Separator,takefocus= 0).grid(row=0, column=9, padx=5, pady=0)
        Label(button_frame, text="Camera Speed setting (1/s):").grid(row=0, column=10, padx=7, pady=5)
        speedCombo = ttk.Combobox(button_frame, textvariable = self.speedSetting, state='readwrite', width=8,  postcommand = self.update_cb_list)
        speedCombo['values']= ['1', '2', '4', '8', '15', '30', '60', '125', '250', '500', '1000', '2000', '4000']
        speedCombo.grid(row=0, column=11, padx=5, pady=5)

        button_frame.pack(expand=False, fill='x')
        
        # Tree (table) definition
        tree = TreeviewEdit(frame, name = "measureTable")
        tree.config(yscrollcommand=v_scroll.set)
        v_scroll.pack(side=RIGHT, fill=Y)

        
        tree['columns'] = [(dataDef.id) for dataDef in self.dataDefs]

        tree.column("#0", width=0,  stretch=NO)
        for dataDef in self.dataDefs:
            tree.column(dataDef.id, anchor=CENTER, width=dataDef.colWidth, stretch = YES)    

        tree.heading("#0",text="",anchor=CENTER)
        for dataDef in self.dataDefs:
            tree.heading(dataDef.id, text=dataDef.colLabel, anchor=CENTER)

        tree.pack(expand=True, fill='both', padx=2, pady=2)
        v_scroll.config(command=tree.yview)

        # start event pump
        self.ws.bind('<<Measure>>', self.dataEvent)

        # TEST : change the thread to start in order to generate test data
        if(self.test):
            Thr=threading.Thread(target=testMeasureThread)
        else:
            Thr=threading.Thread(target=measureThread)
        Thr.start()

        self.ws.mainloop()

if __name__ == "__main__":  

    test = len(sys.argv)>1 and (sys.argv[1] == 'test')
    print("test={}".format(test))

    app = RemoteApp(test=test)
    app.run()
