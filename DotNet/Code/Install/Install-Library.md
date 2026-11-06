# Installing the .NET library (XpressNet.NET)

XpressNet.NET is the .NET version of the XpressNet library. Version 2.0 has the same functionality as version 2.0 of the VB6 DLL, the macOS version and the Linux version. The library can be used in VB.NET and C# applications (Windows Forms), with Visual Studio 2019 or newer and .NET Framework 4.7.2 or newer.

Compared to the VB6 DLL, the .NET library is easier to install: it does not need to be registered (no regsvr32, no administrator rights), and it does not need any ActiveX controls (Winsock, MSComm). The library is a single file, XpressNet.NET.dll, which is placed next to the application (exe).

## Option 1: use the DLL
1. Download [XpressNet.NET.dll](../XpressNet.NET.dll) (and, to be able to step into the library with the debugger, [XpressNet.NET.pdb](../XpressNet.NET.pdb)). Click on the file, and next on the download RAW file icon at the right.
2. Copy the file(s) to a directory of your project, for example a directory named Lib.
3. In Visual Studio: Project → Add Reference… → Browse… and select XpressNet.NET.dll.

When the project is built, Visual Studio copies XpressNet.NET.dll to the directory of the exe (bin\Debug or bin\Release). To install the application on another PC, copy the exe together with XpressNet.NET.dll.

## Option 2: build the library from the source code
The source code is in the directory [XpressNet.NET](../XpressNet.NET). Open XpressNet.NET.sln in Visual Studio, select Release and choose Build → Build Solution. The DLL is created in bin\Release. It is also possible to add the project XpressNet.NET.csproj to the solution of your application (File → Add → Existing Project…), and add a reference to that project (Project → Add Reference… → Projects).

## Using the library
The use of the library is described in the manual of the VB6 DLL (Documentation/HowToUse). In VB.NET, the object is declared and the events are connected as follows:

```vb
Imports XpressNet

Public Class Form1
    Private XpressNet As XpressNetClass

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        XpressNet = New XpressNetClass()
        AddHandler XpressNet.InterfaceConnected, AddressOf InterfaceConnected
        AddHandler XpressNet.Feedback, AddressOf Feedback
        Call XpressNet.CmdConnectTcp("192.168.0.200", "5550")
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        XpressNet.Dispose()                     ' Closes the connection
    End Sub

    Private Sub InterfaceConnected()
        Call XpressNet.CmdAccessoryAuto(12, 1)
    End Sub

    Private Sub Feedback(ByVal address As Integer, ByVal nibble As Byte, ByVal data As Byte)
        Label1.Text = address & ": " & data
    End Sub
End Class
```

Instead of AddHandler, it is also possible to declare the object with WithEvents (`Private WithEvents XpressNet As XpressNetClass`) and to use Handles (`Private Sub XpressNet_Feedback(...) Handles XpressNet.Feedback`).

The events are raised on the thread that created the XpressNetClass object. If the object is created in Form_Load (or in the declaration of the form), the event handlers may therefore directly change the controls of the form; Invoke is not needed.

## New in version 2.0
- After an error answer of the interface the library sends a message again; as long as the command station does not address the interface, sending is paused.
- New events: `SendError(Reason, Description)` for every message that can not be sent, and `InterfaceNichtAdressiert()` / `InterfaceWiederAdressiert()`. The values of Reason are in `SendErrorReasonEnum`.
- `CmdConnectUsb` has the optional parameter `Handshake` (default off).
- By default the logfile is in `C:\Users\Public\TMC-Logfiles`.
- Z21: after `CmdInitUdp` the library sets the broadcast flags (new command `CmdZ21SetBroadcastFlags`).

The details are in the [Reference manual](../../../Documentation/Reference/Reference.md), sections 10 and 11. Because of the extra parameter of CmdConnectUsb, an application that used version 1.3 must be built again; the calls need not change.

## From the VB6 DLL (COM) to the .NET library
A VB.NET application that uses the VB6 DLL via COM can switch to the .NET library as follows:
1. Remove the reference to XpressNet (Solution Explorer → References → XpressNet → Remove).
2. Add a reference to XpressNet.NET.dll (see Option 1).
3. Build the application again.

The names of the methods, events and attributes, and the types of the parameters, are the same as in the VB6 DLL. The calls and event handlers in the application can therefore stay as they are. There are two exceptions:
- The attribute ZentraleVersion is called ZentraleSoftwareVersion. In .NET an event and an attribute can not have the same name; the event ZentraleVersion keeps its name.
- `XpressNet = Nothing` does not close the connection. Use `XpressNet.Dispose()` instead.

## Differences with the VB6 DLL
- Opening a new connection first closes an existing connection.
- As long as the connection is not (yet) established, messages are not sent (logfile and event SendError). In the VB6 DLL a message during the setup of the TCP connection closed the connection.
- A parameter value that does not fit in a byte is not sent (logfile and event SendError). VB6: run time error "Overflow".
- An error (exception) in an event handler of the application does not disturb the library: the error is written to the logfile and afterwards passed to the application in the usual way.
- Calls from another thread are passed to the thread that created the object (a line is written to the logfile).
- If the logfile can not be written, the text is lost without an error (VB6: run time error).

The logfile is the same as for the VB6 DLL: TMC-LogFile.txt in the directory C:\Users\Public\TMC-Logfiles, or in the directory set with LogDirectory.
