//
//  UDPSocket.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  The UDP connection to the Z21 (VB6: UdpClass with the Winsock control).
//  Receiving takes place on a background thread; the received packets are passed via Post to the main thread
//  (see XpressNetMainThread.cs). Sending takes place on the main thread.
//
//  As in VB6, the local port is selected by Windows (bind to port 0). The Z21 spec prefers port 21106 for
//  the client, but with a fixed local port only a single instance of the UDP class would be possible.
//  Note: the Yamorc YD7001D retransmits broadcast messages three times.
//

using System;
using System.Net;
using System.Net.Sockets;

namespace XpressNet
{
  internal class UDPSocketClass
  {
    private readonly XpressNetClass Parent;
    private UdpClient _client;
    private IPEndPoint _remote;
    private volatile bool _closed;                      // close has been called. Set on the main thread
    internal bool ready;                                // Can we send? Main thread only

    internal UDPSocketClass(XpressNetClass parent)
    {
      Parent = parent;
    }


    // ==============================================================================================================
    // Called on the main thread
    internal void init(string Host, string Port)
    {
      int portNumber;
      if (!int.TryParse(Port, out portNumber) || (portNumber < 1) || (portNumber > 65535)) {
        Parent.Logfile.Errors("UDP Error: invalid port: " + Port);
        return;
      }
      try {
        IPAddress address;
        if (!IPAddress.TryParse(Host, out address)) {
          address = null;
          foreach (IPAddress a in Dns.GetHostAddresses(Host)) {
            if (a.AddressFamily == AddressFamily.InterNetwork) { address = a; break; }
          }
          if (address == null) {
            Parent.Logfile.Errors("UDP Error: no IPv4 address for host: " + Host);
            return;
          }
        }
        _remote = new IPEndPoint(address, portNumber);
        _client = new UdpClient(0, AddressFamily.InterNetwork);   // Let Windows select a random free port
        ready = true;
        // 2.0: which connection belongs to the letter of this XpressNet object
        Parent.Logfile.Connection(Host + ", port " + portNumber + ", local port " + ((IPEndPoint)_client.Client.LocalEndPoint).Port);
        _client.BeginReceive(receiveCallback, _client);
      }
      catch (Exception e) {
        ready = false;
        Parent.Logfile.Errors("UDP Error: " + e.Message);
      }
    }

    internal void close()
    {
      _closed = true;
      ready = false;
      try {
        if (_client != null) _client.Close();
      }
      catch (Exception) {
        // The socket is closed anyway
      }
    }

    internal void SendData(byte[] outBuffer)
    {
      if (!ready) return;
      try {
        _client.Send(outBuffer, outBuffer.Length, _remote);
      }
      catch (Exception e) {
        Parent.Logfile.Errors("UDP Error while sending: " + e.Message);
      }
    }


    // ==============================================================================================================
    // Called on a background thread
    private void receiveCallback(IAsyncResult result)
    {
      UdpClient client = (UdpClient)result.AsyncState;
      byte[] udpData = null;
      try {
        IPEndPoint sender = new IPEndPoint(IPAddress.Any, 0);
        udpData = client.EndReceive(result, ref sender);
      }
      catch (ObjectDisposedException) {
        return;                                         // The socket has been closed
      }
      catch (Exception e) {
        // For example: Windows reports an ICMP "port unreachable" of a previous send as an error on the
        // next receive. The socket remains usable, so we continue receiving
        if (_closed) return;
        string description = e.Message;
        Parent.mainThread.Post(() => Parent.Logfile.Details("UDP receive: " + description));
      }
      if (udpData != null) {
        Parent.mainThread.Post(() => { if (!_closed) Parent.receiveZ21.addToIndata(udpData); });
      }
      if (_closed) return;
      try {
        client.BeginReceive(receiveCallback, client);
      }
      catch (Exception e) {
        if (_closed) return;
        string description = e.Message;
        Parent.mainThread.Post(() => Parent.Logfile.Errors("UDP Error while receiving: " + description));
      }
    }
  }
}
