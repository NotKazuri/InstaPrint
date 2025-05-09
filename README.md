# InstaPrint
This is the repository for our Investigatory Project.

**To set the application app:**

1. Create a 2 folders. For the main application folder and 'Select a file' folder (for user's file to be chosen)
2. Now, select the main application folder.
3. Using VS Code's terminal, create a virtual environment for the dependencies  (dependencies are the main support or framework that will help to build this application)
    - In VS Code terminal enter this 'python3 -m venv myenv' and make sure that it's inside the main application folder.
    - Enter this command 'source venv/bin/activate' in order to activate the virtual environment.
4. Now, install the needed dependencies. Enter 'pip install -r requirements.txt'.
5. Now go to this G-Drive to install the interface: https://drive.google.com/drive/folders/1oaKsFycvvnNZheqDd6Gxa27768Z81wfi
    - After clicking the link, go to 'FULL INTERFACE OF INSTAPRINTMACHINE' > Select all the frame from frame 0 to frame5.
    - Extract the folder and put it inside the main application folder.
6. Since everything is done, changed the appropriate file location in the code.
    - Press CRTL + F and search 'InstaPrint' as the keyword and changed the appropriate file location depending on your situation.
7. To set up the printers, press CRTL + F and search long or short. After that, just change the current name of the printer. (Make sure that it's the exact name of each printer)

#Changing the interface
NOTE: This app is fully customizable, depending to your chosen design the developer can change the whole app's interface.
After installing the dependencies and seting up the main application's whole file, now proceed to installing the interface of the app.

The interface can be change, just go to: https://www.figma.com/design/8qyESmi2f3OMeXrn6ikQMK/InstaPrint-Interface?t=aG8ZKlJSRj7vBDfr-1

#To set up the payment system:

The things that are going to be use are:
- Arduino UNO (component)
- Arduino IDE (software)
- Coin Slot (component)
- Power Supply (component)
- 
1. Follow this tutorial: https://www.youtube.com/watch?v=l3SVj6t4sq0w
2. After setting up the Coin Slot, now proceed to programming the Arduino Uno.
    - Copy the script in GitHub for the Coin slot > paste in Arduino IDE.
    - Set up the pins and wirings of the Power Supply and Coin Slot.
    - Make sure to check the proper COM and baud rate in Arduino IDE if they correctly alligned in Python (if not, changed it)
    - Now the Coin Slot should work without any errors.
