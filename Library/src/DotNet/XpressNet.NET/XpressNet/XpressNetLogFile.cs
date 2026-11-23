//
//  XpressNetLogFile.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  .NET version of the LogFileClass of the VB6 XpressNet DLL.
//  Same logfile (TMC-LogFile.txt, with the backups TMC-LogFile-1..3.txt), same directory and same text.
//  The logfile directory can be changed by the application program (LogDirectory)
//  2.0: the default directory is %PUBLIC%\TMC-Logfiles (1.3: %PUBLIC%). On some Windows 11 systems writing directly
//  in %PUBLIC% gives an access error; a subdirectory can be written. The library creates the directory. If that
//  fails, the directory is kept anyway (the logfile can then not be written; %PUBLIC% itself is not used).
//
//  THE FOLLOWING METHODS ARE DEFINED:
//  OpenLogfile                        ' Check the Logfile size, and rotate the Logfiles if needed
//  Append          some Text          ' Append "some text" to the Logfile
//  AppendWithHost  some text          ' Append "some text", preceded by the letter of the XpressNet object
//  Connection      some text          ' 2.0: Append "= some text" (the connection of the XpressNet object)
//  SendHex         byteArray          ' Append "=>" followed by the hex values of byteArray
//  ReceiveHex      byteArray          ' Append "<=" followed by the hex values (without the last byte)
//  ReceiveHexRange byeArray, from, to ' Append "<=" followed by the hex values between from and to
//
//  Details         some Text          ' Same as AppendWithHost, but checks Parent.LogDetails first
//  DetailsSent     some Text          ' 2.0: as Details, preceded by "=>" (description of a sent message)
//  DetailsReceived some Text          ' 2.0: as Details, preceded by "<=" (description of a received message)
//  Errors          some Text          ' Same as AppendWithHost, but checks Parent.LogErrors first
//  Warnings        some Text          ' Same as AppendWithHost, but checks Parent.LogWarnings first
//
//  2.0: every line of an XpressNet object starts with its letter (A, B, ...; see XpressNetClass.instanceId); when a
//  connection is made, a line "A = 192.168.0.200, port 5550, local port 54321" tells which connection the letter
//  belongs to. 1.3: every line ended with the IP address or COM port (myHost, which is still set)
//
//  Differences with the VB6 version:
//  - If the logfile can not be written (for example: no access rights), the text is lost without an error
//    (VB6: run time error)
//

using System;
using System.IO;
using System.Text;

namespace XpressNet
{
  public class LogFileClass
  {
    private readonly XpressNetClass Parent;

    // The maximum Logfile size as well as the names of the Logfile and backups may be modified
    private const long maxLogFileSize = 100000;        // Size is in bytes
    private const string defaultLogSubDirectory = "TMC-Logfiles";   // 2.0: subdirectory of %PUBLIC%

    private string m_LogDir;
    public string logFileName { get; set; }
    public string logFileNameOld1 { get; set; }
    public string logFileNameOld2 { get; set; }
    public string logFileNameOld3 { get; set; }

    public string myHost { get; set; }                  // Allows the Logfile Hex dumps to include the IP address / USB port


    internal LogFileClass(XpressNetClass parent)
    {
      Parent = parent;
      myHost = "";
      // In case the application doesn't provide a path, use the following defaults
      m_LogDir = Environment.GetEnvironmentVariable("PUBLIC");   // Use the Windows Environ variable
      if (string.IsNullOrEmpty(m_LogDir)) {
        m_LogDir = @"C:\Users\Public";                 // Works on several systems, but may require admin rights
      }
      // 2.0: a subdirectory of PUBLIC, created if it does not exist yet (see the header of this file)
      m_LogDir = Path.Combine(m_LogDir, defaultLogSubDirectory);
      try {
        if (!Directory.Exists(m_LogDir)) Directory.CreateDirectory(m_LogDir);
      }
      catch (Exception) {
        // The directory could not be created. The logfile can not be written; the text is lost without an error
      }
      setFileNames();
    }

    private void setFileNames()
    {
      logFileName = Path.Combine(m_LogDir, "TMC-LogFile.txt");
      logFileNameOld1 = Path.Combine(m_LogDir, "TMC-LogFile-1.txt");
      logFileNameOld2 = Path.Combine(m_LogDir, "TMC-LogFile-2.txt");
      logFileNameOld3 = Path.Combine(m_LogDir, "TMC-LogFile-3.txt");
    }


    public void OpenLogfile()
    {
      // The goal of this subroutine is to check the size of the existing Logfile.
      // If that size exceeds maxLogFileSize, the Logfiles will be rotated and renamed.
      try {
        if (File.Exists(logFileName)) {
          long fileSize = new FileInfo(logFileName).Length;
          if (fileSize > maxLogFileSize) {
            // Delete the oldest file and rename the others
            if (File.Exists(logFileNameOld3)) File.Delete(logFileNameOld3);
            if (File.Exists(logFileNameOld2)) File.Move(logFileNameOld2, logFileNameOld3);
            if (File.Exists(logFileNameOld1)) File.Move(logFileNameOld1, logFileNameOld2);
            File.Move(logFileName, logFileNameOld1);
          }
        }
      }
      catch (Exception) {
        // The logfiles could not be rotated. We continue with the existing logfile
      }
      // Write an empty line, to indicate a restart
      writeLine("");
    }

    public void Append(string data)
    {
      writeLine(timeText() + ": " + data);
    }

    public void Details(string data)
    {
      if (Parent.LogDetails) AppendWithHost(data);
    }

    // 2.0: the description of a message gets the same direction marker as its hex values
    internal void DetailsSent(string data)
    {
      Details("=> " + data);
    }

    internal void DetailsReceived(string data)
    {
      Details("<= " + data);
    }

    public void Errors(string data)
    {
      if (Parent.LogErrors) AppendWithHost(data);
    }

    public void Warnings(string data)
    {
      if (Parent.LogWarnings) AppendWithHost(data);
    }

    public void AppendWithHost(string data)
    {
      // 2.0: preceded by the letter of the XpressNet object (1.3: followed by the host)
      // writeLine(timeText() + ": " + data + " (" + myHost + ")");
      writeLine(timeText() + ": " + Parent.instanceId + " " + data);
    }

    // 2.0: the connection of this XpressNet object
    internal void Connection(string data)
    {
      AppendWithHost("= " + data);
    }

    public void SendHex(byte[] byteArray)
    {
      // create a string with Hex values
      StringBuilder hexString = new StringBuilder();
      for (int i = 0; i < byteArray.Length; i++) {
        hexString.Append(byteArray[i].ToString("X")).Append(" ");
      }
      writeLine(timeText() + ": " + Parent.instanceId + " => " + hexString.ToString().TrimEnd());
    }

    public void ReceiveHexRange(byte[] byteArray, byte msgStart, byte msgEnd)
    {
      // create a string with Hex values
      StringBuilder hexString = new StringBuilder();
      for (int i = msgStart; (i <= msgEnd) && (i < byteArray.Length); i++) {
        hexString.Append(byteArray[i].ToString("X")).Append(" ");
      }
      writeLine(timeText() + ": " + Parent.instanceId + " <= " + hexString.ToString().TrimEnd());
    }

    public void ReceiveHex(byte[] byteArray)
    {
      // create a string with Hex values (as in VB6: without the last byte)
      StringBuilder hexString = new StringBuilder();
      for (int i = 0; i < byteArray.Length - 1; i++) {
        hexString.Append(byteArray[i].ToString("X")).Append(" ");
      }
      writeLine(timeText() + ": " + Parent.instanceId + " <= " + hexString.ToString().TrimEnd());
    }

    // The directory for the logfiles. If the directory does not exist, it will be created.
    // If that fails, the previous directory is kept
    public string LogDirectory
    {
      get { return m_LogDir; }
      set {
        if (value == null) return;
        string d = value.Trim();
        if (d.Length == 0) return;
        // remove a trailing backslash
        if (d.EndsWith(@"\")) d = d.Substring(0, d.Length - 1);
        // check folder exists
        if (!Directory.Exists(d)) {
          try {
            Directory.CreateDirectory(d);
          }
          catch (Exception) {
            // Folder creation failed -> keep old m_LogDir and defaults
            return;
          }
        }
        // Once we are here, the folder already exists or has been created
        m_LogDir = d;
        setFileNames();
      }
    }


    // ==============================================================================================================
    // Internal use
    private static string timeText()
    {
      DateTime t = DateTime.Now;
      return t.Year + "/" + t.Month.ToString("00") + "/" + t.Day.ToString("00") + " " + t.Hour.ToString("00") + ":" +
             t.Minute.ToString("00") + ":" + t.Second.ToString("00") + "." + t.Millisecond.ToString("000");
    }

    private void writeLine(string text)
    {
      try {
        File.AppendAllText(logFileName, text + "\r\n", Encoding.Default);
      }
      catch (Exception) {
        // The logfile could not be written. The text is lost
      }
    }
  }
}
