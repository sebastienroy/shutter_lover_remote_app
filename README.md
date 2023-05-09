# shutter_lover_remote_app
The remote application of Shutter Lover tool

## Installation

This application is written in Python language and so can be executed on any platform that supports Python language.  

Here is the installation instructions for Windows platform. Steps are similar on other platforms.  

1. Install Python packages:    
   See https://www.python.org/downloads/
2. Install PySerial package:  
   - On Windows Search Tool, type "cmd" and click on the Command Prompt application.  
   -  Into the command prompt, execute the following line :  
   
   ```
        pip install pyserial
        
   ```
3. Download the file shutter_lover_remote_app.py  

5. Right-click on the downloaded file and select "Open With", then select Python.  
   You can choose to always use Python to open files having ".py" extension, so that you can execute the application through a double-click on the file.

## Usage

- Connect the Shutter Lover tool on your computer using the provided USB cable
- Switch on the tool
- Open the Shutter Lover Remote Application
- The application should automatically discover the communication port used by the tool and establish the communication.
- If not connected, use the combo box to choose the right communication port
- Choose de shutter curtain translation direction, using the dedicated combo box. This choice has an impact on the extrapolated course timing (see above) 
- Perform your shutter measurements. Each time a measurement is made, a new line is added in the Remote Application, containing the full details of the measurement data.
- In order to save your data, click on "Copy to Clipboard" button, then paste the data (CTRL+V) on any external application such as spreadsheet application or text editor.
- To clean the data, click on the "Clear All" button.

## Data Description

| Column Name | Description |
| :--: |:-- |
| Id | The row number. Reset to 0 when clearing the table. |
| Speed (1/s) | The shutter opening speed measured on the central sensor. This is the inverse of the open time. Unit : 1/s |
| Time (ms) | The shutter open time measured on the central sensor. Unit : ms |
| Open (ms) | The course time of the first curtain (opening curtain) of the shutter, measured between the sensors located on the corners of the frame. Unit : ms |
| Close (ms) | The course time of the second curtain (closing curtain) of the shutter, measured between the sensors located on the corners of the frame. Unit : ms |
| Open ext | The sensors of the Shuter Lover are located on the corners of a 20mm x 32mm rectangle. This value is the extrapolation of the first curtain course time on the whole 24x36 film frame. The extrapolation calculation depends on the selected translation direction. Unit: ms |
| Close ext | Extrapolation of the second curtain course time. Unit: ms |




