# Download and Install the XpressNet DLL (VB6)

## Step 1: Create a directory ##

The first step is to open the File Explorer, and create a directory called XpressNetDll. The location and name of the directory should not be changed later. A good place is directly under C:.
<br/><img src="Download/01-Create_Directory.png" alt="Create directory" width="500"/>
<br/><br/>

## Step 2: Go to the GitHub repository ##
Open your browser and go to the [XpressNet repository on Gihub](https://github.com/aikopras/XpressNet).
<br/><img src="Download/02-Goto_GitHub.png" alt="go to GitHub" width="500"/>
<br/><br/>

## Step 3: Double Click VB6_Source.zip ##
Go to the directory called Code and double click VB6_Source.zip
<br/><img src="Download/03-Double_Click.png" alt="click VB6_Source.zip" width="500"/>
<br/><br/>

## Step 4: Download as RAW file ##
Click the download RAW file icon at the right to download VB6_Source.zip to your local machine.
<br/><img src="Download/04-Download_raw_file.png" alt="Download RAW" width="500"/>
<br/><br/>

## Step 5: Open the VB6 Source file ##
After the file is downloaded (to your local Downloads directory), click Open file to Open the VB6_Source file.
<br/><img src="Download/05-Open_VB6_Source.png" alt="Open VB6_Source" width="500"/>
<br/><br/>

## Step 6: Copy files ##
A window opens, showing all the files in the VB6_Source file. Select all files, and copy them to the directory you've created in Step 1 above (like C:/XpressNetDll).
<br/><img src="Download/06-Copy_Files.png" alt="Copy Files" width="600"/>
<br/><br/>

## Step 7: Open VB6 as Admin ##
Now that all files are downloaded and saved, it is time to make and install the DLL at your local machine. For that purpose you have to run VB6 as Administrator.
Click Windows Start -> Microsoft Visual Studio followed by a right click of Microsft Visual Basic 6. A windows appears and under More you select Run as Administrator.
<br/><img src="Make/01-Run_VB6_AsAdmin.png" alt="Open VB6 as Admin" width="600"/>
<br/><br/>

## Step 8: Allow VB6 as Admin ##
Depending on your version of windows, a window may pop-up, asking you to allow VB6 to make changes to your local computer. Click YES.
<br/><img src="Make/02-Allow_VB6_AsAdmin.png" alt="Allow VB6 as Admin" width="400"/>
<br/><br/>

## Step 9: VB6 opens with New Project ##
Depending on your VB6 preferences, a window may pop-up, asking you if you want to create a new VB6 project. Click Cancel.
<br/><img src="Make/03-Cancel_New_Project.png" alt="Cancel new project" width="500"/>
<br/><br/>

## Step 10: Open the XpressNet project ##
Open the XpressNet_vbp VB6 project, which was stored (in C:/XpressNetDll) in Step 6 above.
<br/><img src="Make/04-OpenXpressNet_vbp.png" alt="Open XpressNet.vbp" width="500"/>
<br/><br/>

## Step 11: Make XpressNet.dll ##
Click File -> Make XpressNet.dll, to compile the DLL.
<br/><img src="Make/05-MakeDll.png" alt="Make XpressNet.dll" width="500"/>
<br/><br/>

## Step 12: Save and register the DLL ##
In the last step you are asked where the DLL file should be stored. Select the directory in which you have the source files. Under the hood VB6 also registers the DLL in the Windows Registry, allowing any VB6 program to use the DLL. You are now ready with the installation of the DLL.
<br/><img src="Make/06-SaveDll.png" alt="Save XpressNet.dll" width="500"/>
<br/><br/>

## Step 13: Reference the XpressNet DLL ##
Before you can use the XpressNet DLL library in your project, you must reference the DLL. After opening your project click
Project -> References, and select the XpressNet Interface. See both figures below.
<br/><img src="Reference/01-OpenPreferences.png" alt="Open Project -> Preferences" width="350"/>
<br/><br/>
<br/><img src="Reference/02-SelectXpressNet.png" alt="Select XpressNet" width="500"/>
<br/><br/>
