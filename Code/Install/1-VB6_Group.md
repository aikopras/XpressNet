# Add source code to the application
This manual describes how to use the XpressNet Library source code as part of your VB6 application.


## Step 1: Go to the GitHub repository ##
Open your browser and go to the [XpressNet repository on Gihub](https://github.com/aikopras/XpressNet).
<br/><img src="Download/02-Goto_GitHub.png" alt="go to GitHub" width="500"/>
<br/><br/>

## Step 2: Double Click VB6_Source.zip ##
Go to the directory called Code and double click VB6_Source.zip
<br/><img src="Download/03-Double_Click.png" alt="click VB6_Source.zip" width="500"/>
<br/><br/>

## Step 3: Download as RAW file ##
Click the download RAW file icon at the right to download VB6_Source.zip to your local machine.
<br/><img src="Download/04-Download_raw_file.png" alt="Download RAW" width="500"/>
<br/><br/>

## Step 4: Open the VB6 Source file ##
After the file is downloaded to your local Downloads directory, you must open the received ZIP file. Depending on the browser being used, this may be possible via a pop-up window.
<br/><img src="Download/05-Open_VB6_Source.png" alt="Open VB6_Source" width="500"/>
<br/><br/>

## Stap 5: Copy all files ##
After the ZIP file is opened, a new window opens showing all files contained within VB6_Source.zip. Select all files, and copy them to the folder you want to use for your own VB6 application.
<br/><img src="Download/06-Copy_Files.png" alt="Kopieer files" width="600"/>
<br/><br/>

If your application folder was called "MyApplication", the file explorer shows the following files.
<br/><img src="VB6_Group/01-FileExlplorer.png" alt="Overzicht files" width="400"/>
<br/><br/>

## Step 6: Start a new Standard EXE project ##
To use the XpressNet sources in your own application program, you have to open a new Visual Basic 6.0 project (add project) that will later contain the code for your own application. In general your own application program will take the form of a **Standard EXE** .
<br/><img src="VB6_Group/03-Add_Project.png" alt="Add project" width="250"/>
<br/>
<br/><img src="VB6_Group/04-Standard_exe.png" alt="Standard Exe" width="500"/>
<br/><br/>

## Step 7: Open the XpressNet project and copy the code ##
Open the file named **XpressNet.vbp**, which is the Visual Basic 6 project group that holds together the various XpressNet source files.
<br/><img src="VB6_Group/02-Open_Vbp.png" alt="Open XpressNet.vbp" width="300"/>
<br/><br/>

In your Standard EXE project go to **Project → Add File**

Add all **.frm** and **.cls** files that are present in **XpressNet.vbp**.

You now have all required code directly in your own project. Save your project and close **XpressNet.vbp**.

## Step 8: Compile and run your application ##

Choose **File → Make Project.exe** to compile your program.

You will get a single .exe, containing all code to use XpressNet. Registering an XpressNet DLL is not (any longer) necessary.
