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

