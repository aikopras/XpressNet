//
//  TCPSocket.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  The TCP connection to the Lenz LAN/USB Interface 23151 (VB6: TcpClass with the Winsock control).
//  Connecting and receiving take place on background threads; the results are passed via Post to the main thread
//  (see XpressNetMainThread.cs). Sending takes place on the main thread.
//
//  Events (all on the main thread):
//  - connection established          => Parent.eventInterfaceConnected
//  - connection could not be made    => Parent.eventTcpError (with the description of the error)
//  - connection closed by the remote host, or lost  => Parent.eventTcpClosedByRemoteHost
//  - data received                   => Parent.receive.addToIndata
//
//  2.0: a connection attempt is stopped after 5 seconds (Windows itself gives up after about 21 seconds, for
//  example after a typing error in the IP address). The event TcpError gets a short description (as the
//  Winsock control in VB6); the logfile gets the complete description of the error
//

using System;
using System.Net.Sockets;

namespace XpressNet
{
  internal class TCPSocketClass
  {
    private readonly XpressNetClass Parent;
    private TcpClient _client;
    private NetworkStream _stream;
    private readonly byte[] _readBuffer = new byte[1024];
    private volatile bool _closed;                      // closeConnection has been called. Set on the main thread
    internal bool connected;                            // Is the connection established? Main thread only
    private string _host;                               // 2.0: for the logfile
    private int _port;
    private const int connectTimeOut = 5000;            // 2.0: ms. Maximum time for a connection attempt
    private XTimerClass connectTimer;                   // 2.0: stops the connection attempt after connectTimeOut

    internal TCPSocketClass(XpressNetClass parent)
    {
      Parent = parent;
    }


    // ==============================================================================================================
    // Called on the main thread
    internal void connect(string Host, string Port)
    {
      int portNumber;
      if (!int.TryParse(Port, out portNumber) || (portNumber < 1) || (portNumber > 65535)) {
        Parent.eventTcpError("Invalid TCP port: " + Port);
        return;
      }
      _host = Host;
      _port = portNumber;
      try {
        _client = new TcpClient();
        _client.NoDelay = true;                         // XpressNet messages are small, and should be sent immediately
        connectTimer = new XTimerClass(Parent.mainThread);
        connectTimer.Timer += connectTimer_Timer;
        connectTimer.Interval = connectTimeOut;
        _client.BeginConnect(Host, portNumber, connectCallback, _client);
      }
      catch (Exception e) {
        stopConnectTimer();
        Parent.eventTcpError(shortDescription(e), e.Message);
      }
    }

    // 2.0: the connection attempt takes too long
    private void connectTimer_Timer()
    {
      stopConnectTimer();
      if (_closed || connected) return;
      closeConnection();                                // The result of the connection attempt is ignored
      Parent.eventTcpError("Connection timed out", "Connection attempt timed out after " + (connectTimeOut / 1000) + " seconds");
    }

    private void stopConnectTimer()
    {
      if (connectTimer != null) connectTimer.close();
      connectTimer = null;
    }

    // 2.0: a short description for the event TcpError (the description of .NET is long, and depends on the language
    // of Windows)
    private static string shortDescription(Exception e)
    {
      SocketException s = e as SocketException;
      if (s == null) return e.Message;
      switch (s.SocketErrorCode) {
        case SocketError.ConnectionRefused: return "Connection refused";
        case SocketError.TimedOut: return "Connection timed out";
        case SocketError.HostUnreachable: return "Host unreachable";
        case SocketError.NetworkUnreachable: return "Network unreachable";
        case SocketError.HostNotFound: return "Host not found";
        default: return s.SocketErrorCode.ToString();
      }
    }

    internal void closeConnection()
    {
      _closed = true;
      connected = false;
      stopConnectTimer();
      try {
        if (_client != null) _client.Close();
      }
      catch (Exception) {
        // The connection is closed anyway
      }
    }

    internal void SendData(byte[] outBuffer)
    {
      // STEP 1: Do we still have the TCP connection?
      if (!connected) {
        Parent.eventTcpClosedByRemoteHost();            // Clean up, and inform the application
        return;
      }
      // STEP 2: Send the outbuffer
      try {
        _stream.Write(outBuffer, 0, outBuffer.Length);
      }
      catch (Exception e) {
        Parent.Logfile.Errors("TCP Error while sending: " + e.Message);
        connected = false;
        Parent.eventTcpClosedByRemoteHost();            // Clean up, and inform the application
      }
    }


    // ==============================================================================================================
    // Called on background threads
    private void connectCallback(IAsyncResult result)
    {
      TcpClient client = (TcpClient)result.AsyncState;
      try {
        client.EndConnect(result);
      }
      catch (Exception e) {
        string description = shortDescription(e);
        string logText = e.Message;
        Parent.mainThread.Post(() => { if (!_closed) { stopConnectTimer(); Parent.eventTcpError(description, logText); } });
        return;
      }
      Parent.mainThread.Post(() => {
        if (_closed) return;
        stopConnectTimer();
        try {
          _stream = client.GetStream();
          connected = true;
          // 2.0: which connection belongs to the letter of this XpressNet object
          Parent.Logfile.Connection(_host + ", port " + _port + ", local port " + ((System.Net.IPEndPoint)client.Client.LocalEndPoint).Port);
          _stream.BeginRead(_readBuffer, 0, _readBuffer.Length, readCallback, _stream);
        }
        catch (Exception e) {
          connected = false;
          Parent.eventTcpError(shortDescription(e), e.Message);
          return;
        }
        Parent.eventInterfaceConnected();
      });
    }

    private void readCallback(IAsyncResult result)
    {
      NetworkStream stream = (NetworkStream)result.AsyncState;
      int bytesReceived;
      try {
        bytesReceived = stream.EndRead(result);
      }
      catch (Exception e) {
        string description = e.Message;
        Parent.mainThread.Post(() => connectionLost(description));
        return;
      }
      if (bytesReceived == 0) {                         // The remote host has closed the connection
        Parent.mainThread.Post(() => connectionLost(""));
        return;
      }
      byte[] tcpData = new byte[bytesReceived];         // Copy the TCP segment: _readBuffer is used for the next read
      Array.Copy(_readBuffer, tcpData, bytesReceived);
      Parent.mainThread.Post(() => { if (!_closed) Parent.receive.addToIndata(bytesReceived, tcpData); });
      try {
        stream.BeginRead(_readBuffer, 0, _readBuffer.Length, readCallback, stream);
      }
      catch (Exception e) {
        string description = e.Message;
        Parent.mainThread.Post(() => connectionLost(description));
      }
    }


    // ==============================================================================================================
    // Called on the main thread
    private void connectionLost(string description)
    {
      if (_closed) return;                              // We closed the connection ourselves
      if (description.Length > 0) Parent.Logfile.Errors("TCP Error while receiving: " + description);
      connected = false;
      if (Parent.tcp == this) Parent.eventTcpClosedByRemoteHost();
    }
  }
}
