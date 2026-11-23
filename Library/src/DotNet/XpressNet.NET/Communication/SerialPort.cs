//
//  SerialPort.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  The USB and serial connections (VB6: UsbClass with the MSComm control).
//  Receiving takes place on a background thread; the received bytes are passed via Post to the main thread
//  (see XpressNetMainThread.cs). Sending takes place on the main thread.
//
//  Info regarding Lenz interfaces (from the VB6 version):
//  1) Serial:
//  - LI100: First design. Metal case. Baudrate: 9600.
//  - LI100F: Improved design. Metal case. Baudrates: 9600 (default) and 19200 baud.
//  - LI101F: Last version of the serial interface. Plastic case. Galvanic isolation
//            Baudrates: 19200 (default), 38400, 57600, en 115200 baud.
//  All serial interfaces use:
//  - 8 Databits, 1 Start bit, 1 Stop bit, No parity => Settings = "baudrate,N,8,1"
//  - Hardware-Handshake: the interface uses CTS (Clear To Send) to tell the PC whether or not
//    it is able to receive data. The interface ignores the PC's RTS (Request to Send) signal.
//    Note that the Interface's DSR (Data Set Ready) and the PC's DTR (Data terminal Ready) are not used.
//  2) USB
//  - LI-USB (23150), LI-USB-Ethernet (23151), LZV200, Yamorc YD7001
//  All Lenz USB interfaces use:
//  - 57600 Baud, 8 Databits, 1 Start bit, 1 Stop bit, No parity => Settings = "57600,N,8,1"
//  - No hardware handshake, according to the Lenz documentation (LI-USB 06/2008, Interface 23151 V1.2).
//    2.0: JMRI uses hardware handshake (RTS/CTS) for the LI-USB 23150. Therefore CmdConnectUsb has the optional
//    parameter Handshake (default: off), just like CmdConnectSerial.
//  The Yamorc YD7001 USB interface seems to work with 57600, as well as 115200.
//
//  Received bytes may contain parts of an XpressNet message, or (parts of) multiple XpressNet messages.
//  ReceiveClass.addToIndata takes care of this.
//
//  As in VB6: if the port can not be opened, a MessageBox shows the error.
//

using System;
using System.IO.Ports;
using System.Windows.Forms;

namespace XpressNet
{
  internal class SerialPortClass
  {
    private readonly XpressNetClass Parent;
    private System.IO.Ports.SerialPort _port;
    private volatile bool _closed;                      // closePort has been called. Set on the main thread
    internal bool isOpen;                               // Is the port open? Main thread only

    internal SerialPortClass(XpressNetClass parent)
    {
      Parent = parent;
    }


    // ==============================================================================================================
    // Called on the main thread
    internal void openPort(byte portNumber, string portSettings, bool Handshake)
    {
      try {
        _port = new System.IO.Ports.SerialPort("COM" + portNumber);
        applySettings(_port, portSettings);
        if (Handshake) {
          _port.Handshake = System.IO.Ports.Handshake.RequestToSend;  // Serial uses CTS for hardware handshaking
        }
        else {
          _port.Handshake = System.IO.Ports.Handshake.None;           // No hardware handshaking
          _port.RtsEnable = false;                      // The Lenz XpressNet spec says it ignores RTS
        }
        _port.DtrEnable = false;                        // DSR/DTR is not implemented by Lenz
        _port.WriteTimeout = 2000;                      // Sending should never block the main thread for long
        _port.DataReceived += dataReceived;
        _port.ErrorReceived += errorReceived;
        _port.Open();                                   // Try to open the USB port
        isOpen = true;
        // 2.0: which connection belongs to the letter of this XpressNet object
        Parent.Logfile.Connection("COM" + portNumber + ", " + portSettings + (Handshake ? ", handshake" : ""));
      }
      catch (Exception e) {
        // Catch errors if the port does not exist
        isOpen = false;
        try { MessageBox.Show("COM port:" + portNumber + "\r\n" + " Error: " + e.Message); }
        catch (Exception) { }
        Parent.Logfile.Errors("COM port: " + portNumber + " Error: " + e.Message);
        return;
      }
      Parent.eventInterfaceConnected();
    }

    internal void closePort()
    {
      _closed = true;
      isOpen = false;
      try {
        if ((_port != null) && _port.IsOpen) _port.Close();
      }
      catch (Exception) {
        // The port is closed anyway
      }
    }

    internal void SendData(byte[] outBuffer)
    {
      if (!isOpen) return;
      try {
        _port.Write(outBuffer, 0, outBuffer.Length);
      }
      catch (Exception e) {
        Parent.Logfile.AppendWithHost("COM Error: " + e.Message);
      }
    }

    // Settings as in the MSComm control: "baudrate,parity,databits,stopbits", for example "57600,N,8,1"
    private static void applySettings(System.IO.Ports.SerialPort port, string settings)
    {
      string[] parts = settings.Split(',');
      if (parts.Length != 4) throw new ArgumentException("Invalid settings: " + settings);
      port.BaudRate = int.Parse(parts[0].Trim());
      switch (parts[1].Trim().ToUpperInvariant()) {
        case "N": port.Parity = Parity.None; break;
        case "E": port.Parity = Parity.Even; break;
        case "O": port.Parity = Parity.Odd; break;
        case "M": port.Parity = Parity.Mark; break;
        case "S": port.Parity = Parity.Space; break;
        default: throw new ArgumentException("Invalid parity in settings: " + settings);
      }
      port.DataBits = int.Parse(parts[2].Trim());
      switch (parts[3].Trim()) {
        case "1": port.StopBits = StopBits.One; break;
        case "1.5": port.StopBits = StopBits.OnePointFive; break;
        case "2": port.StopBits = StopBits.Two; break;
        default: throw new ArgumentException("Invalid stop bits in settings: " + settings);
      }
    }


    // ==============================================================================================================
    // Called on a background thread
    private void dataReceived(object sender, SerialDataReceivedEventArgs e)
    {
      try {
        int bytesToRead = _port.BytesToRead;
        if (bytesToRead <= 0) return;
        byte[] usbData = new byte[bytesToRead];
        int bytesReceived = _port.Read(usbData, 0, bytesToRead);
        Parent.mainThread.Post(() => { if (!_closed) Parent.receive.addToIndata(bytesReceived, usbData); });
      }
      catch (Exception) {
        // The port has been closed
      }
    }

    private void errorReceived(object sender, SerialErrorReceivedEventArgs e)
    {
      string text;
      switch (e.EventType) {
        case SerialError.Frame: text = "COM Error: Framing Error"; break;
        case SerialError.Overrun: text = "COM Error: Overrun: Data Lost"; break;
        case SerialError.RXOver: text = "COM Error: RX Overflow - Receive buffer overflow"; break;
        case SerialError.RXParity: text = "COM Error: Parity Error"; break;
        case SerialError.TXFull: text = "COM Error: TX buffer full"; break;
        default: text = "COM Error: " + e.EventType; break;
      }
      Parent.mainThread.Post(() => { if (!_closed) Parent.Logfile.AppendWithHost(text); });
    }
  }
}
