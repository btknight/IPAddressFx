using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace IPAddressFx
{
    [Guid("6D4A28CE-0792-4C02-8778-031A0DE4A0A5")]
    [ComVisible(true)]
    public interface IIPAddressFX
    {
        object IPAndNetmask4ToCIDRHost(string Address, string Bitmask);
        object IPAndNetmask4ToCIDRNetwork(string Address, string Bitmask);
        object IPAndNetmask4ToNetwork(string Address, string Bitmask);
        object ConvertMaskToBitlen(string Bitmask);
        object ConvertBitlenToMask4(decimal Bitlen);
        object IsInSameNetwork4(string Address, string Network, [Optional] object Netmask);
        object Inet_aton(string Address);
        object Inet_ntoa(double AddrAsInteger);
        object Summarize4(Excel.Range AddrRange);
        object GetHostAddresses(string Hostname);
        object GetHostName(string Address);
        object FindASN(string Address);
        object Ping(string Address, [Optional] object NumPings, [Optional] object Timeout);
        object TCPConnect(string Address, decimal Port, [Optional] object Timeout);
    }
    
    /// <summary>
    /// Contains useful functions to manipulate IP addresses within Excel spreadsheets.
    /// </summary>
    [Guid("3B4CEFC2-E170-4BD6-B827-F18EE1238CB4")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IIPAddressFX))]
    public class IPAddressFx : IIPAddressFX
    {
        public IPAddressFx()
        {
        }

        /// <summary>
        /// Converts an IP address string and a netmask to VLSM slash-style notation (ww.xx.yy.zz/nn).
        /// </summary>
        /// <param name="Address">IP address</param>
        /// <param name="Bitmask">Network mask</param>
        /// <returns>IP address in VLSM slash-style notation</returns>
        public object IPAndNetmask4ToCIDRHost(string Address, string Bitmask)
        {
            IPAddress _address;
            if(!IPAddress.TryParse(Address, out _address))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            object bitlen = ConvertMaskToBitlen(Bitmask);
            if (bitlen is ErrorWrapper)
            {
                return bitlen;
            }
            return String.Format("{0}/{1}", _address.ToString(), bitlen);
        }

        /// <summary>
        /// Takes an IP address string and a netmask string, and returns the VLSM notation for the 
        /// network block where the IP address resides.
        /// </summary>
        /// <param name="Address">IP address in dotted-quad format</param>
        /// <param name="Bitmask">Network mask in dotted-quad format</param>
        /// <returns>Network block, in VLSM slash-style notation</returns>
        public object IPAndNetmask4ToCIDRNetwork(string Address, string Bitmask)
        {
            object Network = IPAndNetmask4ToNetwork(Address, Bitmask);
            if (Network is ErrorWrapper)
            {
                return Network;
            }
            object bitlen = ConvertMaskToBitlen(Bitmask);
            if (bitlen is ErrorWrapper)
            {
                return bitlen;
            }
            return String.Format("{0}/{1}", Network, bitlen);
        }

        /// <summary>
        /// Given an IP address and a network mask, returns the network block to which the IP belongs.
        /// </summary>
        /// <param name="Address">IP address in dotted-quad format</param>
        /// <param name="Bitmask">Network mask in dotted-quad format</param>
        /// <returns>IP address of the network block in dotted-quad format</returns>
        public object IPAndNetmask4ToNetwork(string Address, string Bitmask)
        {
            IPAddress _address;
            if (!IPAddress.TryParse(Address, out _address))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            IPAddress _bitmask;
            if (!IPAddress.TryParse(Bitmask, out _bitmask))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            IPAddress _network;
            try
            {
                _network = _PerformBitmask(_address, _bitmask);
            }
            catch (ArgumentException)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            return _network.ToString();
        }

        /// <summary>
        /// Given a network mask in dotted-quad format, returns the bit length of the mask 
        /// (for VLSM-style notation - the "nn" in ww.xx.yy.zz/nn).
        /// </summary>
        /// <param name="Bitmask">Network mask in dotted-quad format</param>
        /// <returns>Bit length</returns>
        public object ConvertMaskToBitlen(string Bitmask)
        {
            IPAddress addr;
            if (!IPAddress.TryParse(Bitmask, out addr))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            decimal BitLen = 0;

            try
            {
                BitLen = _ConvertMaskToBitlen(addr);
            }
            catch (ArgumentException)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            return BitLen;
        }

        /// <summary>
        /// Given a network mask, returns the bit length of the mask (for VLSM-style notation - the "nn" 
        /// in ww.xx.yy.zz/nn).
        /// </summary>
        /// <param name="Bitmask">IPAddress object containing the network mask</param>
        /// <returns>Bit length</returns>
        private int _ConvertMaskToBitlen(IPAddress Bitmask)
        {
            byte[] bytes = Bitmask.GetAddressBytes();

            int BitLen = 0;
            bool CurrentBitIsOne = false;
            for (int i = bytes.Length - 1; i >= 0; i--)
            {
                for (int j = 0; j < 8; j++)
                {
                    if ((bytes[i] & 1) == 1)
                    {
                        CurrentBitIsOne = true;
                        BitLen++;
                    }
                    else
                    {
                        if (CurrentBitIsOne == true)
                        {
                            throw new ArgumentException("Value is an invalid bitmask");
                        }
                    }
                    bytes[i] >>= 1;
                }
            }
            return BitLen;
        }

        /// <summary>
        /// Given a VLSM bit length (the "nn" in ww.xx.yy.zz/nn), returns the full network mask.
        /// </summary>
        /// <param name="Bitlen">Bit length, between 0 and 32</param>
        /// <returns>Network mask in dotted-quad format</returns>
        public object ConvertBitlenToMask4(decimal Bitlen)
        {
            int _bitlen = (int)Bitlen;
            IPAddress _addr;
            try
            {
                _addr = _ConvertBitlenToMask4(_bitlen);
            }
            catch (ArgumentException)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            return _addr.ToString();
        }

        /// <summary>
        /// Given a VLSM bit length (the "nn" in ww.xx.yy.zz/nn), returns the full network mask.
        /// </summary>
        /// <param name="Bitlen">Bit length, between 0 and 32</param>
        /// <returns>IPAddress object containing the network mask</returns>
        private IPAddress _ConvertBitlenToMask4(int Bitlen)
        {
            if (Bitlen < 0 || Bitlen > 32)
            {
                throw new ArgumentException("Bitlen should be between 0 and 32");
            }
            long addr = 0;
            for (int i = 0; i < 32; i++)
            {
                addr <<= 1;
                addr += (Bitlen > 0) ? 1 : 0;
                Bitlen--;
            }
            byte[] bytes = new byte[4];
            for (int i = 3; i >= 0; i--)
            {
                bytes[i] = (byte)(addr % 256);
                addr >>= 8;
            }
            return new IPAddress(bytes);
        }

        /// <summary>
        /// Given an IP address and a network specification, returns a boolean indicating whether the single 
        /// IP address exists within the network.
        /// </summary>
        /// <param name="Address1">First IP address to test</param>
        /// <param name="Address2">Second IP address to test (can be a network block in VLSM slash-style notation)</param>
        /// <param name="Netmask">Netmask (required if the network block is not in VLSM format)</param>
        /// <returns></returns>
        public object IsInSameNetwork4(string Address1, string Address2, [Optional] object Netmask)
        {
            IPAddress _address1 = null;
            if (!IPAddress.TryParse(Address1, out _address1))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            IPAddress _address2 = null;
            IPAddress _netmask = null;
            if (Address2.IndexOf('/') > 0)
            {
                string[] _networkArg = Address2.Split('/');
                if (!IPAddress.TryParse(_networkArg[0], out _address2))
                {
                    return new ErrorWrapper((int)CVErrNum.ErrValue);
                }
                try
                {
                    _netmask = _ConvertBitlenToMask4(Convert.ToInt32(_networkArg[1]));
                }
                catch (ArgumentException)
                {
                    return new ErrorWrapper((int)CVErrNum.ErrValue);
                }
            }
            else
            {
                if (!IPAddress.TryParse(Address2, out _address2))
                {
                    return new ErrorWrapper((int)CVErrNum.ErrValue);
                }
            }
            if (!(Netmask is System.Reflection.Missing))
            {
                string _netmaskStr = "";
                if (Netmask is string)
                {
                    _netmaskStr = (string)Netmask;
                }
                if (Netmask is Excel.Range)
                {
                    Excel.Range _netmaskRng = (Excel.Range)Netmask;
                    _netmaskStr = (string)_netmaskRng.Value;
                }
                if (!IPAddress.TryParse((string)_netmaskStr, out _netmask))
                {
                    return new ErrorWrapper((int)CVErrNum.ErrValue);
                }
            }
            if (_netmask == null)
            {
                return new ErrorWrapper((int)CVErrNum.ErrNull);
            }
            IPAddress _addr1Masked = _PerformBitmask(_address1, _netmask);
            IPAddress _addr2Masked = _PerformBitmask(_address2, _netmask);
            if (_addr1Masked.Equals(_addr2Masked))
            {
                return true;
            }
            return false;
        }


        /// <summary>
        /// The Excel analogue of the traditional C function "inet_aton".  Converts a dotted-quad
        /// IP address to a numeral.  Useful for sorting.
        /// </summary>
        /// <param name="Address">IP address in dotted-quad format</param>
        /// <returns>IP address in numerical format</returns>
        public object Inet_aton(string Address)
        {
            IPAddress addr;
            if (!IPAddress.TryParse(Address, out addr))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            long retval = 0;
            for (int i = 0; i < addr.GetAddressBytes().Length; i++)
            {
                retval <<= 8;
                retval += (long)addr.GetAddressBytes()[i];
            }
            return (double)retval;
        }

        /// <summary>
        /// The Excel analogue of the traditional C function "inet_ntoa".  Converts a numerical
        /// IP address to a dotted-quad.
        /// </summary>
        /// <param name="Address">IP address in numerical format</param>
        /// <returns>IP address in dotted-quad format</returns>
        public object Inet_ntoa(double AddrAsInteger)
        {
            byte[] _bytes = new byte[4];
            long _addrAsInteger = (long)AddrAsInteger;
            for (int i = 3; i >= 0; i--)
            {
                _bytes[i] = (byte)(_addrAsInteger % 256);
                _addrAsInteger >>= 8;
            }
            IPAddress _addr = new IPAddress(_bytes);
            return _addr.ToString();
        }

        /// <summary>
        /// Given a range of dotted-quad IP addresses, finds a single subnet that encapsulates all 
        /// of the addresses listed.
        /// </summary>
        /// <param name="AddrRange">Range of IP addresses in dotted-quad format</param>
        /// <returns>Network block, in VLSM slash-style format</returns>
        public object Summarize4(Excel.Range AddrRange)
        {
            IPAddress _min = IPAddress.Any;
            IPAddress _max = IPAddress.Any;

            AddressFamily AF = AddressFamily.InterNetwork;
            IPAddressComparer AC = new IPAddressComparer();
            object Values = AddrRange.Value;
            if (AddrRange.Rows.Count == 1 && AddrRange.Columns.Count == 1)
            {
                IPAddress _addr;
                if (!(IPAddress.TryParse(Values.ToString(), out _addr)))
                {
                    return new ErrorWrapper((int)CVErrNum.ErrValue);
                }
                _min = _addr;
                _max = _addr;
            }
            else 
            {
                object[,] _array = (object[,])Values;
                for (int r = 1; r <= _array.GetLength(0); r++)
                {
                    for (int c = 1; c <= _array.GetLength(1); c++)
                    {
                        object cellObj = _array.GetValue(r, c);
                        string cellValue = cellObj.ToString();
                        IPAddress _addr = null;
                        if (!(IPAddress.TryParse(cellValue, out _addr)))
                        {
                            return new ErrorWrapper((int)CVErrNum.ErrValue);
                        }
                        if (AF == AddressFamily.Unknown)
                        {
                            AF = _addr.AddressFamily;
                        }
                        else
                        {
                            if (AF != _addr.AddressFamily)
                            {
                                return new ErrorWrapper((int)CVErrNum.ErrValue);
                            }
                        }
                        if (_min == IPAddress.Any || AC.Compare(_min, _addr) < 0)
                        {
                            _min = _addr;
                        }
                        if (_max == IPAddress.Any || AC.Compare(_max, _addr) > 0)
                        {
                            _max = _addr;
                        }
                    }
                }
            }

            for (int i = 32; i >= 0; i--)
            {
                IPAddress mask = _ConvertBitlenToMask4(i);
                IPAddress _network = _PerformBitmask(_min, mask);
                if (_network.Equals(_PerformBitmask(_max, mask)))
                {
                    return String.Format("{0}/{1}", _network, i);
                }
            }
            // Flow of the program should never get here.
            return "0.0.0.0/0";
        }

        /// <summary>
        /// Given a DNS name, attempts to resolve the hostname to one or more IP addresses.
        /// </summary>
        /// <param name="Hostname">Hostname to look up</param>
        /// <returns>A list of IP addresses</returns>
        public object GetHostAddresses(string Hostname)
        {
            if (Hostname == null || Hostname == "")
            {
                return "";
            }
            IPAddress[] addrs;
            try {
                addrs = Dns.GetHostAddresses(Hostname);
            }
            catch (ArgumentException)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            return String.Join(", ", addrs.Select((x) => { return x.ToString(); }));
        }

        /// <summary>
        /// Given an IP address, attempts to resolve the address to a hostname.
        /// </summary>
        /// <param name="Address">IP address to look up</param>
        /// <returns>Hostname associated with the IP in reverse DNS</returns>
        public object GetHostName(string Address)
        {
            IPAddress addr;
            if (!IPAddress.TryParse(Address, out addr))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            IPHostEntry host;
            try
            {
                host = Dns.GetHostEntry(addr);
            }
            catch (ArgumentException)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            return host.HostName;
        }

        /// <summary>
        /// Given a single IP address, pings the target.  Returns status in the cell.
        /// </summary>
        /// <param name="Address">IP address to ping, in dotted-quad format</param>
        /// <param name="NumPings">Number of times to ping the target (default 3)</param>
        /// <param name="Timeout">Length of time to wait for reply, in milliseconds (default 1000ms)</param>
        /// <returns>Short line describing results</returns>
        public object Ping(string Address, [Optional] object NumPings, [Optional] object Timeout)
        {
            IPAddress addr;
            if (!IPAddress.TryParse(Address, out addr))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            int _numPings = 3;
            if (!(NumPings is System.Reflection.Missing))
            {
                _numPings = Convert.ToInt32(NumPings);
            }
            int _timeout = 1000;
            if (!(Timeout is System.Reflection.Missing))
            {
                _timeout = Convert.ToInt32(Timeout);
            }

            Ping pinger = new Ping();

            List<PingReply> replies = new List<PingReply>();

            int Success = 0;
            for (int i = 0; i < _numPings; i++)
            {
                PingReply reply = pinger.Send(Address, _timeout);
                if (reply.Status == IPStatus.Success) { Success++; }
                replies.Add(reply);
            }

            int RTT = 0;
            if (Success > 0)
            {
                foreach (PingReply reply in replies)
                {
                    if (reply.Status == IPStatus.Success)
                    {
                        RTT += (int)reply.RoundtripTime;
                    }
                }
                RTT /= Success;
                return String.Format("Success {0}/{1}, RTT {2}", Success, replies.Count, RTT);
            }
            else
            {
                return "No reply";
            }
        }

        /// <summary>
        /// Attempts a TCP connection to the given IP address and port.
        /// </summary>
        /// <param name="Address">IP address of the target</param>
        /// <param name="Port">Target TCP port</param>
        /// <param name="Timeout">Timeout in milliseconds. (Defaults to 3000ms.)</param>
        /// <returns></returns>
        public object TCPConnect(string Address, decimal Port, [Optional] object Timeout)
        {
            IPAddress _addr;
            if (!IPAddress.TryParse(Address, out _addr))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            int _port = (int)Port;
            if (_port < 0 || _port > 65535)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            int _timeout = 3000;
            if(!(Timeout is System.Reflection.Missing))
            {
                if (Timeout is decimal || Timeout is double)
                {
                    _timeout = Convert.ToInt32(Timeout);
                }
                if (Timeout is Excel.Range)
                {
                    Excel.Range _timeoutRng = (Excel.Range)Timeout;
                    _timeout = Convert.ToInt32(_timeoutRng.Value);
                }
            }
            if (_timeout < 1)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }

            IPEndPoint _ep = new IPEndPoint(_addr, _port);
            Socket _sock = new Socket(_addr.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            _sock.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Linger, new LingerOption(true, 0));
            //_sock.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.DontLinger, true);
            _sock.Blocking = false;
            bool _success = false;
            DateTime _tZero = DateTime.Now;
            try
            {
                IAsyncResult _ar = _sock.BeginConnect(_ep, null, null);
                _ar.AsyncWaitHandle.WaitOne(_timeout, true);
                _success = _sock.Connected;
                if (_success)
                {
                    _sock.EndConnect(_ar);
                }
            }
            catch (SocketException e)
            {
                if (e.ErrorCode == 10013)
                {
                    return "Failed";
                }
                return String.Format("Cannot connect: {0}: {1}", e.ErrorCode, e.Message);
            }
            TimeSpan _timeTaken = DateTime.Now - _tZero;
            //_sock.Close(0);
            if (_success)
            {
                return String.Format("Success: {0}ms", _timeTaken.TotalMilliseconds);
            }
            else
            {
                return "Failed";
            }
        }


        // http://www.team-cymru.org/Services/ip-to-asn.html#dns
        /// <summary>
        /// Given an IP address, finds the origin AS number as advertised in BGP.  Uses cymru.com DNS services to map IP to ASN.
        /// </summary>
        /// <param name="Address">IP address</param>
        /// <returns>Bar-delimited field showing AS number, component prefix, countrry of origin, and date of ASN registration with RIR</returns>
        public object FindASN(string Address)
        {
            IPAddress _addr;
            if (!IPAddress.TryParse(Address, out _addr))
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            if (_addr.AddressFamily != AddressFamily.InterNetwork && _addr.AddressFamily != AddressFamily.InterNetworkV6)
            {
                return new ErrorWrapper((int)CVErrNum.ErrValue);
            }
            StringBuilder sb = new StringBuilder();
            List<string> AddrBytesStrings = new List<string>();
            byte[] AddrBytes = _addr.GetAddressBytes();
            bool FoundNonZeroValue = false;
            for(int i = AddrBytes.Length - 1; i >= 0; i--)
            {
                if(AddrBytes[i] > 0 || (AddrBytes[i] == 0 && FoundNonZeroValue == true))
                {
                    FoundNonZeroValue = true;
                    if(_addr.AddressFamily == AddressFamily.InterNetworkV6)
                    {
                        AddrBytesStrings.Add(String.Format("{0:x1}", AddrBytes[i] % 0x10));
                        AddrBytesStrings.Add(String.Format("{0:x1}", AddrBytes[i] / 0x10));
                    }
                    else
                    {
                        AddrBytesStrings.Add(String.Format("{0}", AddrBytes[i]));
                    }
                }
            }
            sb.Append(String.Join(".", AddrBytesStrings));
            if (_addr.AddressFamily == AddressFamily.InterNetworkV6)
            {
                sb.Append(".origin6.asn.cymru.com");
            }
            else
            {
                sb.Append(".origin.asn.cymru.com");
            }
            string[] output;
            try
            {
                output = DnsapiQuery.GetTXTRecord(sb.ToString());
            }
            catch (Win32Exception e)
            {
                return e.Message;
            }
            return output[0];
        }
        
        
        /// <summary>
        /// Performs a masking operation (bitwise AND) on two IPAddress objects.
        /// </summary>
        /// <param name="Addr1">IPAddress object for the first IP address</param>
        /// <param name="Addr2">IPAddress object for the second IP address</param>
        /// <returns>New IPAddress object containing the masked IP address</returns>
        protected static IPAddress _PerformBitmask(IPAddress Addr1, IPAddress Addr2)
        {
            if (Addr1.AddressFamily != Addr2.AddressFamily)
            {
                throw new ArgumentException("Addr1 and Addr2 address families are not the same - cannot mask");
            }
            int AddrLen = Addr1.GetAddressBytes().Length;
            byte[] _addrMaskedBytes = new byte[AddrLen];
            for (int i = 0; i < AddrLen; i++)
            {
                _addrMaskedBytes[i] = (byte)(Addr1.GetAddressBytes()[i] & Addr2.GetAddressBytes()[i]);
            }
            return new IPAddress(_addrMaskedBytes);
        }

        // http://xldennis.wordpress.com/2006/11/29/dealing-with-cverr-values-in-net-part-ii-solutions/
        /// <summary>
        /// Internal function to test whether the incoming argument from Excel is an error or not.
        /// Throws an ArgumentException if the argument presented is invalid.
        /// </summary>
        /// <param name="obj">Boxed object to test</param>
        /// <returns>True if the object is an error, false if it is a normal argument</returns>
        protected static bool IsXLCVErr(object obj)
        {
            if (obj is int)
            {
                switch ((int)obj)
                {
                    case (int)CVErrNum.ErrDiv0:
                        return true;
                    case (int)CVErrNum.ErrNA:
                        return true;
                    case (int)CVErrNum.ErrName:
                        return true;
                    case (int)CVErrNum.ErrNull:
                        return true;
                    case (int)CVErrNum.ErrNum:
                        return true;
                    case (int)CVErrNum.ErrRef:
                        return true;
                    case (int)CVErrNum.ErrValue:
                        return true;
                    default:
                        throw new ArgumentException("Object passed in is an int32 and from a non-COM source.");
                }
            }
            if (obj is double || obj is string || obj is decimal || obj is DateTime || obj is bool)
            {
                return false;
            }
            if (obj == null)
            {
                return false;
            }
            if (obj is Excel.Range)
            {
                throw new ArgumentException("Excel.Range object passed.  Pass Range.Value instead");
            }
            if (obj is Array)
            {
                throw new ArgumentException("Array object passed.  Method valid only for single-cell values");
            }
            throw new ArgumentException("Unknown argument type " + obj.GetType().ToString());
        }

        // http://blogs.msdn.com/b/eric_carter/archive/2004/12/01/273127.aspx
        /// <summary>
        /// Registers this object as a COM object.
        /// </summary>
        /// <param name="type"></param>
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(
              GetSubKeyName(type, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(
              GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("",
              System.Environment.SystemDirectory + @"\mscoree.dll",
              RegistryValueKind.String);
        }

        /// <summary>
        /// Unregisters this object as a COM object.
        /// </summary>
        /// <param name="type"></param>
        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(
              GetSubKeyName(type, "Programmable"), false);
        }

        /// <summary>
        /// Used by COM registration functions to build pathnames for subkeys in the registry.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="subKeyName"></param>
        /// <returns></returns>
        private static string GetSubKeyName(Type type,
          string subKeyName)
        {
            System.Text.StringBuilder s =
              new System.Text.StringBuilder();
            s.Append(@"CLSID\{");
            s.Append(type.GUID.ToString().ToUpper());
            s.Append(@"}\");
            s.Append(subKeyName);
            return s.ToString();
        }  

    }

    /// <summary>
    /// Compares two byte arrays for sorting.
    /// </summary>
    public class AddrByteComparer : Comparer<byte[]>
    {
        public override int Compare(byte[] x, byte[] y)
        {
            if (x.Length != y.Length)
            {
                return x.Length.CompareTo(y.Length);
            }
            for (int i = 0; i < x.Length; i++)
            {
                if (x[i] != y[i])
                {
                    return x[i].CompareTo(y[i]);
                }
            }
            return 0;
        }
    }

    /// <summary>
    /// Compares two IPAddress objects for sorting.
    /// </summary>
    public class IPAddressComparer : Comparer<IPAddress>
    {
        public override int Compare(IPAddress x, IPAddress y)
        {
            AddrByteComparer ABC = new AddrByteComparer();
            return ABC.Compare(x.GetAddressBytes(), y.GetAddressBytes());
        }
    }

    /// <summary>
    /// Enumeration containing Excel errors as integers.
    /// </summary>
    public enum CVErrNum : int
    {
        /// <summary>
        /// Divide by zero
        /// </summary>
        ErrDiv0 = -2146826281,
        
        /// <summary>
        /// Not available / not applicable
        /// </summary>
        ErrNA = -2146826246,

        /// <summary>
        /// Function name invalid
        /// </summary>
        ErrName = -2146826259,
        
        /// <summary>
        /// Value is null
        /// </summary>
        ErrNull = -2146826288,

        /// <summary>
        /// Numerical computation error
        /// </summary>
        ErrNum = -2146826252,

        /// <summary>
        /// Reference error
        /// </summary>
        ErrRef = -2146826265,

        /// <summary>
        /// Value is invalid
        /// </summary>
        ErrValue = -2146826273,
    }
}
