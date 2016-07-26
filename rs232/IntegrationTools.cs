using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;
using System.EnterpriseServices;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Net.NetworkInformation;
using System.Drawing;
using System.Drawing.Imaging;
using System.Data.SqlClient;
using System.Diagnostics;

namespace AP.IntegrationTools
{
    #region Rs232
    [Guid("99116080-9EA1-4C2F-A001-380674478086"), 
        ComVisible(true), 
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRs232
    {
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        string Help();
        /// <summary>
        /// Фцнкція зважування.
        /// </summary>
        /// <param name="port">COM порт з яконо потрібно отримати вагу (вказується повністю COM1,COM31)</param>
        /// <param name="cpuType">Тип вагопроцессору (доступні значення: 1-Техноваги"Залізничні"; 2-Техноваги"Авто"; 3-Булат)</param>
        /// <param name="cpu">Номер вагопроцессора в лінії (зазвичай 1)</param>
        /// <returns>Повертає ціле числове значення - вагу в кілограмах.</returns>
        int GetWeight(string port, int cpuType, int cpu);
        /// <summary>
        /// Функція переводу порта з підключеним зчитувачем Z2USB в режим прослуховування.
        /// </summary>
        /// <param name="port">COM порт до якого підключений зчитувач Z2USB</param>
        void OpenZ2usb(string port);
    }

    [Guid("99116080-9EA1-4C2F-E001-380674478086"), 
        ComVisible(true), 
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRs232Events
    {
        /// <summary>
        /// Зовнішня подія, яка повертає код карти при зчитуванні з порту.
        /// </summary>
        /// <param name="rfid">Текстове поле з кодом зчитаної карти.</param>
        [DispId(1)]
        void OnDataRead(string rfid);
    }


    [Guid("99116080-9EA1-4C2F-C001-380674478086"),
        ComVisible(true),
        ClassInterface(ClassInterfaceType.None),
        ComSourceInterfaces(typeof(IRs232Events))]
    public class Rs232 : IRs232
    {
        #region Internal Helpers
        [ComVisible(false)]
        public delegate void OnMyEvtDelgate(string rfid);

        [DispId(1)]
        public event OnMyEvtDelgate OnDataRead;

        public Rs232()
        {

        }


        private void Dummy(string rfid)
        {
            //File.WriteAllText("c:\\deploy\\z2read.txt", String.Format("{0}:{1}", DateTime.Now.ToString(), rfid));
        }

        private void OnSerialPortDataRecived(object sender, SerialDataReceivedEventArgs e)
        {
            string data = ((SerialPort)sender).ReadLine();
            //OnDataRecieved("Test");
            //OnDataRecieved(((SerialPort)sender).ReadLine().ToString());
            try
            {
                OnDataRead("Test");
            }
            catch (Exception ex)
            {
                File.AppendAllText("c:\\deploy\\z2read.txt", ex.Message + Environment.NewLine);
            }
            //File.AppendAllText("c:\\deploy\\z2read.txt", String.Format("{0} :-: {1}\r\n", DateTime.Now.ToString(), data));

        }

        //Техноваги Вагон
        private int GetWeightFromTVV(string port)
        {
            SerialPort _serialPort = new SerialPort(port, 19200, Parity.None, 8, StopBits.One);
            _serialPort.PortName = port;
            try
            {
                _serialPort.Open();
                _serialPort.Write(new byte[] { 0x00 }, 0, 1);
                Thread.Sleep(200);
                if (_serialPort.BytesToRead == 0)
                {
                    _serialPort.Close();
                    throw new ApplicationException("Відсутня відповідь від вагопроцесора!");
                }
                byte[] buffer = new byte[_serialPort.BytesToRead];
                _serialPort.Read(buffer, 0, buffer.Length);
                _serialPort.Close();
                StringBuilder sb = new StringBuilder();
                foreach (byte b in buffer)
                    sb.AppendFormat("{0:X2} ", b.ToString("X2"));
                string checksum = (buffer[0] + buffer[1] + buffer[2] + buffer[3] + buffer[4] + buffer[5] + buffer[6] + buffer[7] + buffer[8] + 1).ToString("X2");
                if (checksum.Substring(checksum.Length - 2, 2) == buffer[9].ToString("X2"))
                {
                    int weight = Convert.ToInt32(buffer[2].ToString("X2") + buffer[1].ToString("X2"), 16);
                    weight += Convert.ToInt32(buffer[4].ToString("X2") + buffer[3].ToString("X2"), 16);
                    weight += Convert.ToInt32(buffer[6].ToString("X2") + buffer[5].ToString("X2"), 16);
                    weight += Convert.ToInt32(buffer[8].ToString("X2") + buffer[7].ToString("X2"), 16);
                    return weight * 10;
                }
                else
                {
                    throw new ApplicationException("Невірна контрольна сума!");
                }
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                throw ex;
            }
        }

        //Техноваги Старі 
        private int GetWeightFromWE2110Print(string port)
        {
            SerialPort _serialPort = new SerialPort(port, 9600, Parity.None, 8, StopBits.One);
            _serialPort.ReadTimeout = 5000;
            string result = "";
            try
            {
                _serialPort.Open();
                for (int i = 0; i < 10; i++)
                {
                    Thread.Sleep(100);
                    try
                    {
                        result = _serialPort.ReadExisting();
                        if (result.IndexOf('-') == -1)
                        {
                            if (result.IndexOf('G') > -1)
                            {
                                try
                                {
                                    int weight = Convert.ToInt32(result.Substring(result.IndexOf(' ') + 1, 7).Replace(".", "").Replace(" ", "")) * 10;
                                    _serialPort.Close();
                                    return weight;
                                }
                                catch { }
                            }
                            else if (result.IndexOfAny(new char[] { 'G', 'M', 'N', 'O', 'E', 'U' }) == -1)
                            {
                                try
                                {
                                    int weight = Convert.ToInt32(result.Substring(result.IndexOf(' ') + 1, 7).Replace(".", "").Replace(" ", "")) * 10;
                                    _serialPort.Close();
                                    return weight;
                                }
                                catch { }
                            }

                        }
                        else
                        {
                            if (result.IndexOf('G') > -1)
                            {
                                try
                                {
                                    int weight = Convert.ToInt32(result.Substring(result.IndexOf('-') + 1, 7).Replace(".", "").Replace(" ", "")) * 10;
                                    _serialPort.Close();
                                    return weight;
                                }
                                catch { }
                            }
                            else if (result.IndexOfAny(new char[] { 'G', 'M', 'N', 'O', 'E', 'U' }) == -1)
                            {
                                try
                                {
                                    int weight = Convert.ToInt32(result.Substring(result.IndexOf('-') + 1, 7).Replace(".", "").Replace(" ", "")) * 10;
                                    _serialPort.Close();
                                    return weight;
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
                _serialPort.Close();
                throw new ApplicationException("Не вдалось отримати вагу за 10 спроб!");
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                throw ex;
            }
        }

        //Техноваги Старі
        private int GetWeightFromWE2110Network(string port, int cpu)
        {
            SerialPort _serialPort = new SerialPort(port, 9600, Parity.None, 8, StopBits.One);
            string result;
            try
            {
                _serialPort.ReadTimeout = 5000;
                _serialPort.Open();
                _serialPort.Write("S" + string.Format("{0:d2}", cpu) + ";\r\n");
                _serialPort.Write("MSV?;\r\n");
                Thread.Sleep(200);
                try
                {
                    result = _serialPort.ReadLine();
                }
                catch
                {
                    _serialPort.Close();
                    throw new ApplicationException("Відсутня відповідь від вагопроцесора!");
                }
                try
                {
                    return Convert.ToInt32(result.Substring(0, 7).Trim().Replace(",", "").Replace(".", "")) * 10;
                }
                catch (Exception ex1)
                {
                    throw ex1;
                }
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                throw ex;
            }
        }

        //Техноваги ТВП (Авто)
        private int GetWeightFromTVP(string port, int cpu)
        {
            SerialPort _serialPort = new SerialPort(port, 9600, Parity.None, 8, StopBits.One);
            _serialPort.ReadTimeout = 5000;
            string result;
            try
            {
                _serialPort.Open();
                for (int i = 0; i < 10; i++)
                {
                    _serialPort.Write("G" + string.Format("{0:d2}", cpu) + "W\r\n");
                    Thread.Sleep(50);
                    try
                    {
                        result = _serialPort.ReadLine();
                    }
                    catch
                    {
                        _serialPort.Close();
                        throw new ApplicationException("Відсутня відповідь від вагопроцесора №" + string.Format("{0:d2}", cpu) + "!");
                    }
                    if (result.Substring(0, 4) == "R" + string.Format("{0:d2}", cpu) + "W")
                    {
                        try
                        {
                            int weight = Convert.ToInt32(result.Substring(4, (result.Length - 4)).Replace(".", String.Empty).Replace(",", String.Empty)) * 10;
                            _serialPort.Close();
                            return weight;
                        }
                        catch { }
                    }
                }
                _serialPort.Close();
                throw new ApplicationException("Не вдалось отримати вагу за 10 спроб!");
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                throw ex;
            }
        }

        //Булат
        private int GetWeightFromIT3000D(string port)
        {
            SerialPort _serialPort = new SerialPort(port, 9600, Parity.None, 8, StopBits.One);
            string result;
            try
            {
                _serialPort.ReadTimeout = 5000;
                _serialPort.Open();
                _serialPort.Write("<RN>\r\n");
                Thread.Sleep(200);
                try
                {
                    result = _serialPort.ReadLine();
                }
                catch
                {
                    _serialPort.Close();
                    throw new ApplicationException("Відсутня відповідь від вагопроцесора!");
                }
                try
                {
                    int weight = Convert.ToInt32(result.Substring(23, 8).Trim());
                    _serialPort.Close();
                    return weight;
                }
                catch (Exception ex1)
                {
                    throw ex1;
                }
            }
            catch (Exception ex)
            {
                _serialPort.Close();
                throw ex;
            }
        }
        #endregion

        #region Public Methotds
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        public string Help()
        {
            MethodInfo[] methodInfos = this.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
            StringBuilder methods = new StringBuilder();
            methods.AppendLine(AssemblyAbout.GetVersion());
            foreach (MethodInfo method in methodInfos)
            {
                methods.Append(method.ReturnType + " " + method.Name + "(");
                if (method.GetParameters().Length > 0)
                {
                    foreach (ParameterInfo parameter in method.GetParameters())
                    {
                        methods.Append(parameter.ParameterType + " " + parameter.Name + ",");
                    }
                    methods.Remove(methods.Length - 1, 1);
                }
                methods.AppendLine(")");
            }
            return methods.ToString();
        }
        /// <summary>
        /// Функція переводу порта з підключеним зчитувачем Z2USB в режим прослуховування.
        /// </summary>
        /// <param name="port">COM порт до якого підключений зчитувач Z2USB</param>
        public void OpenZ2usb(string port)
        {
            SerialPort _serialPort = new SerialPort(port, 9600, Parity.None, 8, StopBits.One);
            _serialPort.DataReceived += new SerialDataReceivedEventHandler(OnSerialPortDataRecived);
            //this.OnDataRead += new OnMyEvtDelgate(Dummy);
            try
            {
                _serialPort.Open();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// Фцнкція зважування.
        /// </summary>
        /// <param name="port">COM порт з яконо потрібно отримати вагу (вказується повністю COM1,COM31)</param>
        /// <param name="cpuType">Тип вагопроцессору (доступні значення: 1-Техноваги"Залізничні"; 2-Техноваги"Авто"; 3-Булат)</param>
        /// <param name="cpu">Номер вагопроцессора в лінії (зазвичай 1)</param>
        /// <returns>Повертає ціле числове значення - вагу в кілограмах.</returns>
        public int GetWeight(string port, int cpuType, int cpu)
        {
            switch (cpuType)
            {
                case 1:
                    try
                    {
                        return (GetWeightFromTVV(port));
                    }
                    catch (Exception ex)
                    {
                        if (ex.InnerException != null)
                            throw ex.InnerException;
                        else
                            throw ex;
                    }
                case 2:
                    try
                    {
                        return (GetWeightFromTVP(port, cpu));
                    }
                    catch (Exception ex)
                    {
                        if (ex.InnerException != null)
                            throw ex.InnerException;
                        else
                            throw ex;
                    }
                case 3:
                    try
                    {
                        return (GetWeightFromIT3000D(port));
                    }
                    catch (Exception ex)
                    {
                        if (ex.InnerException != null)
                            throw ex.InnerException;
                        else
                            throw ex;
                    }
                case 4:
                    try
                    {
                        return (GetWeightFromWE2110Network(port, cpu));
                    }
                    catch (Exception ex)
                    {
                        if (ex.InnerException != null)
                            throw ex.InnerException;
                        else
                            throw ex;
                    }

                case 5:
                    try
                    {
                        return (GetWeightFromWE2110Print(port));
                    }
                    catch (Exception ex)
                    {
                        if (ex.InnerException != null)
                            throw ex.InnerException;
                        else
                            throw ex;
                    }
                default:
                    throw new ApplicationException("Необхідно вказати тип вагопроцесору!");
            }

        }
        #endregion
    }
    #endregion

    #region Moxa
    [Guid("99116080-9EA1-4C2F-A002-380674478086"), ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMoxaIO
    {
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        string Help();
        /// <summary>
        /// Функція перевірки стану цифрових входів
        /// </summary>
        /// <param name="ip">IP-адреса пристрою</param>
        /// <param name="port">Порт пристрою</param>
        /// <param name="password">Пароль для підключення</param>
        /// <param name="timeout">Таймаут на підключення (мілісекунди)</param>
        /// <param name="di">Кількість входів починаючи з 0</param>
        /// <returns>Повертає масив логічних змінних довжиною, яку передано у параметрі di</returns>
        bool[] GetDIStatus(string ip, ushort port, int timeout, string password, byte di);
    }

    [Guid("99116080-9EA1-4C2F-C002-380674478086"),
        ComVisible(true),
        ClassInterface(ClassInterfaceType.None)]
    public class MoxaIO : IMoxaIO
    {
        #region MOXA helpers

        internal const UInt16 Port = 502;						//Modbus TCP port
        internal const UInt16 DO_SAFE_MODE_VALUE_OFF = 0;
        internal const UInt16 DO_SAFE_MODE_VALUE_ON = 1;
        internal const UInt16 DO_SAFE_MODE_VALUE_HOLD_LAST = 2;

        internal const UInt16 DI_DIRECTION_DI_MODE = 0;
        internal const UInt16 DI_DIRECTION_COUNT_MODE = 1;
        internal const UInt16 DO_DIRECTION_DO_MODE = 0;
        internal const UInt16 DO_DIRECTION_PULSE_MODE = 1;

        internal const UInt16 TRIGGER_TYPE_LO_2_HI = 0;
        internal const UInt16 TRIGGER_TYPE_HI_2_LO = 1;
        internal const UInt16 TRIGGER_TYPE_BOTH = 2;
        //A-OPC Server response W5340 Device STATUS information data filed index
        internal const int IP_INDEX = 0;
        internal const int MAC_INDEX = 4;

        private static string CheckErr(int iRet, string szFunctionName)
        {

            string szErrMsg = "MXIO_OK";

            if (iRet != MOXA_CSharp_MXIO.MXIO_CS.MXIO_OK)
            {

                switch (iRet)
                {
                    case MOXA_CSharp_MXIO.MXIO_CS.ILLEGAL_FUNCTION:
                        szErrMsg = "ILLEGAL_FUNCTION";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.ILLEGAL_DATA_ADDRESS:
                        szErrMsg = "ILLEGAL_DATA_ADDRESS";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.ILLEGAL_DATA_VALUE:
                        szErrMsg = "ILLEGAL_DATA_VALUE";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SLAVE_DEVICE_FAILURE:
                        szErrMsg = "SLAVE_DEVICE_FAILURE";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SLAVE_DEVICE_BUSY:
                        szErrMsg = "SLAVE_DEVICE_BUSY";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.EIO_TIME_OUT:
                        szErrMsg = "EIO_TIME_OUT";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.EIO_INIT_SOCKETS_FAIL:
                        szErrMsg = "EIO_INIT_SOCKETS_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.EIO_CREATING_SOCKET_ERROR:
                        szErrMsg = "EIO_CREATING_SOCKET_ERROR";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.EIO_RESPONSE_BAD:
                        szErrMsg = "EIO_RESPONSE_BAD";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.EIO_SOCKET_DISCONNECT:
                        szErrMsg = "EIO_SOCKET_DISCONNECT";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.PROTOCOL_TYPE_ERROR:
                        szErrMsg = "PROTOCOL_TYPE_ERROR";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_OPEN_FAIL:
                        szErrMsg = "SIO_OPEN_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_TIME_OUT:
                        szErrMsg = "SIO_TIME_OUT";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_CLOSE_FAIL:
                        szErrMsg = "SIO_CLOSE_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_PURGE_COMM_FAIL:
                        szErrMsg = "SIO_PURGE_COMM_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_FLUSH_FILE_BUFFERS_FAIL:
                        szErrMsg = "SIO_FLUSH_FILE_BUFFERS_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_GET_COMM_STATE_FAIL:
                        szErrMsg = "SIO_GET_COMM_STATE_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_SET_COMM_STATE_FAIL:
                        szErrMsg = "SIO_SET_COMM_STATE_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_SETUP_COMM_FAIL:
                        szErrMsg = "SIO_SETUP_COMM_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_SET_COMM_TIME_OUT_FAIL:
                        szErrMsg = "SIO_SET_COMM_TIME_OUT_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_CLEAR_COMM_FAIL:
                        szErrMsg = "SIO_CLEAR_COMM_FAIL";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_RESPONSE_BAD:
                        szErrMsg = "SIO_RESPONSE_BAD";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SIO_TRANSMISSION_MODE_ERROR:
                        szErrMsg = "SIO_TRANSMISSION_MODE_ERROR";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.PRODUCT_NOT_SUPPORT:
                        szErrMsg = "PRODUCT_NOT_SUPPORT";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.HANDLE_ERROR:
                        szErrMsg = "HANDLE_ERROR";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.SLOT_OUT_OF_RANGE:
                        szErrMsg = "SLOT_OUT_OF_RANGE";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.CHANNEL_OUT_OF_RANGE:
                        szErrMsg = "CHANNEL_OUT_OF_RANGE";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.COIL_TYPE_ERROR:
                        szErrMsg = "COIL_TYPE_ERROR";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.REGISTER_TYPE_ERROR:
                        szErrMsg = "REGISTER_TYPE_ERROR";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.FUNCTION_NOT_SUPPORT:
                        szErrMsg = "FUNCTION_NOT_SUPPORT";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.OUTPUT_VALUE_OUT_OF_RANGE:
                        szErrMsg = "OUTPUT_VALUE_OUT_OF_RANGE";
                        break;
                    case MOXA_CSharp_MXIO.MXIO_CS.INPUT_VALUE_OUT_OF_RANGE:
                        szErrMsg = "INPUT_VALUE_OUT_OF_RANGE";
                        break;
                }

                Console.WriteLine("Function \"{0}\" execution Fail. Error Message : {1}\n", szFunctionName, szErrMsg);

                if (iRet == MOXA_CSharp_MXIO.MXIO_CS.EIO_TIME_OUT || iRet == MOXA_CSharp_MXIO.MXIO_CS.HANDLE_ERROR)
                {
                    //To terminates use of the socket
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                }
            }
            return szErrMsg;
        }
        #endregion
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        public string Help()
        {
            MethodInfo[] methodInfos = this.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
            StringBuilder methods = new StringBuilder();
            methods.AppendLine(AssemblyAbout.GetVersion());
            foreach (MethodInfo method in methodInfos)
            {
                methods.Append(method.ReturnType + " " + method.Name + "(");
                if (method.GetParameters().Length > 0)
                {
                    foreach (ParameterInfo parameter in method.GetParameters())
                    {
                        methods.Append(parameter.ParameterType + " " + parameter.Name + ",");
                    }
                    methods.Remove(methods.Length - 1, 1);
                }
                methods.AppendLine(")");
            }
            return methods.ToString();
        }
        /// <summary>
        /// Функція перевірки стану цифрових входів
        /// </summary>
        /// <param name="ip">IP-адреса пристрою</param>
        /// <param name="port">Порт пристрою</param>
        /// <param name="password">Пароль для підключення</param>
        /// <param name="timeout">Таймаут на підключення (мілісекунди)</param>
        /// <param name="di">Кількість входів починаючи з 0</param>
        /// <returns>Повертає масив логічних змінних довжиною, яку передано у параметрі di</returns>
        public bool[] GetDIStatus(string ip, ushort port, int timeout, string password, byte di)
        {
            try
            {
                IPAddress.Parse(ip);
                bool[] result = new bool[di];
                int ret;
                Int32[] hConnection = new Int32[1];
                byte[] bytCheckStatus = new byte[1];
                string error;
                ret = MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Init();
                error = CheckErr(ret, "MXEIO_Init");
                if (!error.Contains("MXIO_OK"))
                {
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                    throw new Exception("MXEIO_Init error: " + error);
                }
                ret = MOXA_CSharp_MXIO.MXIO_CS.MXEIO_E1K_Connect(UTF8Encoding.UTF8.GetBytes(ip), port, (uint)timeout, hConnection, UTF8Encoding.UTF8.GetBytes(password));
                error = CheckErr(ret, "MXEIO_E1K_Connect");
                if (!error.Contains("MXIO_OK"))
                {
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                    throw new Exception("MXEIO_E1K_Connect: " + error);
                }
                ret = MOXA_CSharp_MXIO.MXIO_CS.MXEIO_CheckConnection(hConnection[0], 5000, bytCheckStatus);
                error = CheckErr(ret, "MXEIO_CheckConnection");
                if (!error.Contains("MXIO_OK"))
                {
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                    throw new Exception("MXEIO_E1K_Connect: " + error + "\r\n");
                }
                else if (bytCheckStatus[0] != MOXA_CSharp_MXIO.MXIO_CS.CHECK_CONNECTION_OK)
                {
                    switch (bytCheckStatus[0])
                    {
                        case MOXA_CSharp_MXIO.MXIO_CS.CHECK_CONNECTION_OK:
                            MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                            throw new Exception("MXEIO_CheckConnection: Check connection ok");
                        case MOXA_CSharp_MXIO.MXIO_CS.CHECK_CONNECTION_FAIL:
                            MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                            throw new Exception("MXEIO_CheckConnection: Check connection fail");
                        case MOXA_CSharp_MXIO.MXIO_CS.CHECK_CONNECTION_TIME_OUT:
                            MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                            throw new Exception("MXEIO_CheckConnection: Check connection time out");
                        default:
                            MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                            throw new Exception("XEIO_CheckConnection: Check connection status unknown");
                    }
                }
                UInt32[] dwGetDIValue = new UInt32[1];
                UInt16[] wSetDI_DIMode = new UInt16[di];
                for (int i = 0; i < di; i++)
                    wSetDI_DIMode[i] = DI_DIRECTION_DI_MODE;
                ret = MOXA_CSharp_MXIO.MXIO_CS.E1K_DI_SetModes(hConnection[0], 0, di, wSetDI_DIMode);
                error = CheckErr(ret, "E1K_DI_SetModes");
                if (!error.Contains("MXIO_OK"))
                {
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Disconnect(hConnection[0]);
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                    throw new Exception("MXEIO_E1K_Connect: " + error + "\r\n");
                }

                UInt16[] wFilter = new UInt16[di];
                for (int i = 0; i < di; i++)
                    wFilter[i] = (UInt16)(100);
                ret = MOXA_CSharp_MXIO.MXIO_CS.E1K_DI_SetFilters(hConnection[0], 0, di, wFilter);
                if (!error.Contains("MXIO_OK"))
                {
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Disconnect(hConnection[0]);
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                    throw new Exception("E1K_DI_SetFilters: " + error + "\r\n");
                }
                Thread.Sleep(100);
                ret = MOXA_CSharp_MXIO.MXIO_CS.E1K_DI_Reads(hConnection[0], 0, di, dwGetDIValue);
                if (!error.Contains("MXIO_OK"))
                {
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Disconnect(hConnection[0]);
                    MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
                    throw new Exception("E1K_DI_Reads: " + error + "\r\n");
                }
                for (int i = 0, dwShiftValue = 0; i < di; i++, dwShiftValue++)
                {
                    result[i] = !((dwGetDIValue[0] & (1 << dwShiftValue)) == 0);
                    //richTextBox1.AppendText(String.Format("DI vlaue: ch[{0}] = {1}", i + bytStartChannel, ((dwGetDIValue[0] & (1 << dwShiftValue)) == 0) ? "OFF" : "ON") + Environment.NewLine);
                }
                MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Disconnect(hConnection[0]);
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                MOXA_CSharp_MXIO.MXIO_CS.MXEIO_Exit();
            }
        }
    }

    #endregion

    #region Infratec

    [Guid("99116080-9EA1-4C2F-A003-380674478086"), 
        ComVisible(true), 
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IInfratec
    {
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        string Help();
        /// <summary>
        /// Функція повертає повну інформацію, щодо останнього аналізу, проведенного в момент отримання
        /// </summary>
        /// <param name="infratecIp">IP-адреса аналізатору</param>
        /// <param name="port">Порт на якому очікувати прийом інформаціі</param>
        /// <param name="labTimeout">Таймаут очікування інформації від аналізатору (в секундах)</param>
        /// <param name="pingTimeout">Таймаут очікування відповіді від аналізатору (в мілісекундах, якщо 0 то ігнорується перевірка зв’язку з аналізатором)</param>
        /// <returns>Повертаю строку з вмістом повного тексту аналізу</returns>
        string GetGrainAnalyze(string infratecIp, int port, int labTimeout, int pingTimeout);
        /// <summary>
        /// Функція фільтрує повний результат аналізу від аналізатору
        /// </summary>
        /// <param name="analyze">Строка повного результату аналізу</param>
        /// <returns>Повертає масив строк в форматі Показник=Значення</returns>
        string[] ParseGrainAnalyze(string analyze);
    }


    [Guid("99116080-9EA1-4C2F-C003-380674478086"),
        ComVisible(true),
        ClassInterface(ClassInterfaceType.None)]
    public class Infratec : IInfratec
    {
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        public string Help()
        {
            MethodInfo[] methodInfos = this.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
            StringBuilder methods = new StringBuilder();
            methods.AppendLine(AssemblyAbout.GetVersion());
            foreach (MethodInfo method in methodInfos)
            {
                methods.Append(method.ReturnType + " " + method.Name + "(");
                if (method.GetParameters().Length > 0)
                {
                    foreach (ParameterInfo parameter in method.GetParameters())
                    {
                        methods.Append(parameter.ParameterType + " " + parameter.Name + ",");
                    }
                    methods.Remove(methods.Length - 1, 1);
                }
                methods.AppendLine(")");
            }
            return methods.ToString();
        }
        /// <summary>
        /// Функція повертає повну інформацію, щодо останнього аналізу, проведенного в момент отримання
        /// </summary>
        /// <param name="infratecIp">IP-адреса аналізатору</param>
        /// <param name="port">Порт на якому очікувати прийом інформаціі</param>
        /// <param name="labTimeout">Таймаут очікування інформації від аналізатору (в секундах)</param>
        /// <param name="pingTimeout">Таймаут очікування відповіді від аналізатору (в мілісекундах, якщо 0 то ігнорується перевірка зв’язку з аналізатором)</param>
        /// <returns>Повертаю строку з вмістом повного тексту аналізу</returns>
        public string GetGrainAnalyze(string infratecIp, int port, int labTimeout, int pingTimeout)
        {
            Ping pingSender = new Ping();
            PingOptions options = new PingOptions();
            try
            {
                IPAddress.Parse(infratecIp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            // Use the default Ttl value which is 128,
            // but change the fragmentation behavior.
            options.DontFragment = true;

            // Create a buffer of 32 bytes of data to be transmitted.
            string pingData = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            byte[] buffer = Encoding.ASCII.GetBytes(pingData);
            PingReply reply = null;
            if (pingTimeout > 0)
                reply = pingSender.Send(infratecIp, pingTimeout, buffer, options);
            if (pingTimeout == 0 || reply.Status == IPStatus.Success)
            {
                TcpListener infratecListner = new TcpListener(IPAddress.Any, port);
                string data = String.Empty;
                try
                {

                    infratecListner.Start();
                    Byte[] bytes = new Byte[256];
                    for (int current = 0; current < labTimeout * 10; current++)
                    {
                        Thread.Sleep(100);
                        //Application.DoEvents();
                        if (infratecListner.Pending())
                        {
                            TcpClient infratecClient = infratecListner.AcceptTcpClient();
                            //throw new Exception((((IPEndPoint)infratecClient.Client.RemoteEndPoint).Address).ToString() + " - " + (((IPEndPoint)infratecClient.Client.RemoteEndPoint).Address).ToString());
                            if (infratecIp == (((IPEndPoint)infratecClient.Client.RemoteEndPoint).Address).ToString())
                            {
                                //infratecClient.ReceiveTimeout = 10000;
                                NetworkStream dataStream = infratecClient.GetStream();
                                int i;

                                // Loop to receive all the data sent by the client.
                                while ((i = dataStream.Read(bytes, 0, bytes.Length)) != 0)
                                {
                                    // Translate data bytes to a ASCII string.
                                    data += System.Text.Encoding.UTF8.GetString(bytes, 0, i);
                                    //Console.WriteLine("Received: {0}", data);
                                    //File.AppendAllText("log.txt", data);
                                }
                            }
                            infratecClient.Close();
                        }
                        else if (data.Length > 0)
                            break;
                    }
                    infratecListner.Stop();
                    //Thread.Sleep(1000);
                }
                catch (Exception ex)
                {
                    infratecListner.Stop();
                    throw ex;
                }
                finally
                {
                    infratecListner.Stop();
                }
                return data;
            }
            else
            {
                throw new Exception("Ping Timeout (" + infratecIp + ")");
            }
        }
        /// <summary>
        /// Функція фільтрує повний результат аналізу від аналізатору
        /// </summary>
        /// <param name="analyze">Строка повного результату аналізу</param>
        /// <returns>Повертає масив строк в форматі Показник=Значення</returns>
        public string[] ParseGrainAnalyze(string analyze)
        {
            string criterions = String.Empty;
            string results = String.Empty;
            foreach (var line in analyze.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (line.LastIndexOf("Name=String,") > -1)
                {
                    criterions += line.Substring(line.LastIndexOf("Name=String,") + 12).Replace("\"", String.Empty) + ";";
                }
                if (line.LastIndexOf("/PredictedValue=Number,") > -1)
                {
                    results += line.Substring(line.LastIndexOf("/PredictedValue=Number,") + 23) + ";";
                }
            }
            string[] criterionsArray = criterions.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string[] resultsArray = results.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string[] result = new string[criterionsArray.Length];
            for (int count = 0; count < criterionsArray.Length; count++)
            {
                if (!String.IsNullOrEmpty(criterionsArray[count]) && !String.IsNullOrWhiteSpace(criterionsArray[count])
                    && !String.IsNullOrEmpty(resultsArray[count]) && !String.IsNullOrWhiteSpace(resultsArray[count]))
                {
                    result[count] = criterionsArray[count] + "=" + resultsArray[count];
                }
            }
            return result;
        }
    }

    #endregion

    #region IIDK

    [Guid("99116080-9EA1-4C2F-A004-380674478086"), 
        ComVisible(true), 
        InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface Iiidk
    {
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        string Help();
        /// <summary>
        /// Метод ініціалізації компоненти інтеграції систем відеонагляду та 1Сv8
        /// </summary>
        /// <param name="ip">IP-адреса сервеа відеонагляду (string)</param>
        /// <param name="cameras">Двомірний масив строк {№камери, повний шлях до файлу без розширення} (string[,])</param>
        /// <param name="title">Субтитри, що накладаються на зображення (для переносу строк викристовуться символи "\r\n") (string)</param>
        /// <param name="weightTime">Період часу з якого потрібно отримати зображення (якщо пустий то в реальному часі)</param>
        /// <value>Повертає істину або хибність отримання зображення</value>
        bool SaveImage(string ip, string[,] cameras, string title, string weightTime);
        /// <summary>
        /// Метод для зміни параметрів за замовченням
        /// </summary>
        /// <param name="count">Кількість спроб (int32)</param>
        /// <param name="compression">Розмір компрессії (0 - без зтиснення, 5 - максимальне стиснення) (int32)</param>
        /// <param name="imgformat">Формат збереження зображень (jpg, gif, png, tiff, bmp) (string)</param>
        /// <param name="debug">Режим відладки</param>
        void SetExportParams(int count, int compression, string imgformat, bool debug);
        /// <summary>
        /// Метод для зміни облікових данних за замовченням
        /// </summary>
        /// <param name="login">Користувач (string)</param>
        /// <param name="password">Пароль (string)</param>
        void SetCerdentials(string login, string password);
        /// <summary>
        /// Метод для зміни параметрів титрів
        /// </summary>
        /// <param name="font">Назва шрифту</param>
        /// <param name="size">Розмір шрифту</param>
        /// <param name="textcolor">Кольор шрифту</param>
        /// <param name="strokecolor">Кольор обведення</param>
        /// <param name="strokesize">Розмір обведення</param>
        void SetTitleOptions(string font, int size, string textcolor, string strokecolor, int strokesize);
        /// <summary>
        /// Метод повертає список системних кольорів
        /// </summary>
        /// <returns>string[] масив назв кольорів</returns>
        string[] GetColors();
        /// <summary>
        /// Метод для маркування легальності проїзду транспорту
        /// </summary>
        /// <param name="ip">IP-адреса серверу відеоспостереження</param>
        /// <param name="database">Назва бази данних</param>
        /// <param name="id">Номер ТТН</param>
        /// <param name="carID">Номер автомобіля</param>
        /// <param name="entry">Номер проїзду</param>
        /// <param name="culture">Культура</param>
        /// <param name="arm">Етап бізнес-процесу</param>
        void SendLegal(string ip, string database, string id, string carID, int entry, string culture, string arm);
    }


    [Guid("99116080-9EA1-4C2F-C004-380674478086"),
        ComVisible(true),
        ClassInterface(ClassInterfaceType.None)]
    public class iidk : Iiidk
    {
        #region variables
        string id;
        string cctvIP;
        string titles;
        bool result;
        int current;
        string[,] exportdata;
        string[] camerainfo;
        AxIIDK_COMLib.AxIIDK_COM ocx;
        string exportTime;
        DateTime exportDateTime;
        #endregion

        #region Internal helper methods for IIDK
        bool CheckParams()
        {
            if (Options.Debug)
            {
                Log.Write("AP.IntegrationTools.iidk", "CheckParams", Assembly.GetExecutingAssembly().GetName().ToString());
                Log.Write("AP.IntegrationTools.iidk", "CheckParams", "Перевірка параметрів...");
            }
            IPAddress ip = null;
            if (IPAddress.TryParse(cctvIP, out ip))
            {
                int k = 0;
                for (int i = 0; i < exportdata.GetLength(0); i++)
                {
                    //MessageBox.Show("index:" + i + " data: " + exportdata[i, 0]);
                    if (!(Int32.TryParse(exportdata[i, 0], out k)))
                    {
                        Log.Write("AP.IntegrationTools.iidk", "CheckParams", "Неможливо перетворити строку " + exportdata[i, 0] + " до типу int32!");
                        return false;
                    }
                }
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "CheckParams", "Параметри задовільні."); // + exportdata.GetLength(0) + ", " + exportdata.GetLength(1) + "]");
                return true;
            }
            if (Options.Debug)
                Log.Write("AP.IntegrationTools.iidk", "CheckParams", "Неможливо перетворити строку " + cctvIP + " до типу IP-адреса!");
            return false;
        }

        void ocx_OnVideoFrame(object sender, AxIIDK_COMLib._DIIDK_COMEvents_OnVideoFrameEvent e)
        {
            string command;
            if (current < exportdata.GetLength(0))
            {
                if (e.channel == Int32.Parse(exportdata[current, 0]))
                {

                    if (Options.Debug)
                        Log.Write("AP.IntegrationTools.iidk", "OnVideoFrame", "Збереження зображення у файл " + Path.GetTempPath() + id + e.channel + ".bmp " + GetDateTime(e.sysTime).ToString());
                    ocx.SaveToBitmap(e.pImage, Path.GetTempPath() + id + e.channel + ".bmp");
                    camerainfo[current] = GetDateTime(e.sysTime).ToString();
                    current++;
                    if (current < exportdata.GetLength(0))
                    {
                        if (DateTime.TryParse(exportTime, out exportDateTime))
                        {
                            Log.Write("AP.IntegrationTools.iidk", "OnVideoFrame", "Вдала обробка часу: " + exportDateTime.ToString());
                            command = "CAM|" + exportdata[current, 0] + "|ARCH_FRAME_TIME|time<" + exportDateTime.ToString() + ">";
                        }
                        else
                        {
                            Log.Write("AP.IntegrationTools.iidk", "OnVideoFrame", "LIVE START");
                            command = "CAM|" + exportdata[current, 0] + "|START_VIDEO|compress<" + Options.Compression + ">";
                        }
                        ocx.SendMsg(command);
                        Log.Write("AP.IntegrationTools.iidk", "OnVideoFrame", command);
                    }
                }
            }
            else
            {
                ocx.OnVideoFrame -= ocx_OnVideoFrame;
            }
        }

        static DateTime GetDateTime(int sysTime)
        {
            IntPtr pSysTime = new IntPtr(sysTime);

            SystemTime sTime = new SystemTime();
            sTime = (SystemTime)Marshal.PtrToStructure(pSysTime, typeof(SystemTime));

            return sTime.ToDateTime();
        }

        void ocx_OnConnect(object sender, EventArgs e)
        {
            try
            {
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "OnConnect", "Вдале підключення до " + cctvIP);
                Thread cctvLive = new Thread(() => ocxShow(ocx));
                cctvLive.Start();
            }
            catch (Exception ex)
            {
                Log.Write("AP.IntegrationTools.iidk", "OnConnect", ex.Message);
            }
        }

        void ocxShow(object sender)
        {
            string command;
            Thread.Sleep(100);
            try
            {
                current = 0;
                if (DateTime.TryParse(exportTime, out exportDateTime))
                {
                    command = "CAM|" + exportdata[0, 0] + "|ARCH_FRAME_TIME|time<" + exportDateTime.ToString() + ">";
                }
                else
                {
                    Log.Write("AP.IntegrationTools.iidk", "OnConnect", "LIVE START");
                    command = "CAM|" + exportdata[0, 0] + "|START_VIDEO|compress<" + Options.Compression + ">";
                }
                ocx.SendMsg(command);
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "OnConnect", "Активація зображення з відеокамер...");
            }
            catch (Exception ex)
            {
                Log.Write("AP.IntegrationTools.iidk", "OnConnect", ex.Message);
            }
            ocx.OnConnect -= ocx_OnConnect;
        }

        void CheckFile(object state)
        {
            AutoResetEvent are = (AutoResetEvent)state;
            if (Options.Debug)
                Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Запуск механізму очікування результата...");
            for (int i = 1; i < Options.Count; i++)
            {
                if (result)
                    break;
                Thread.Sleep(1000);
                int count = 0;
                for (int j = 0; j < exportdata.GetLength(0); j++)
                    if (File.Exists(Path.GetTempPath() + id + exportdata[j, 0] + ".bmp"))
                    {
                        ocx.SendMsg("CAM|" + exportdata[j, 0] + "|STOP_VIDEO");
                        count++;
                    }
                #region if all files exported
                if (count == exportdata.GetLength(0))
                {
                    ocx.OnVideoFrame -= ocx_OnVideoFrame;
                    if (Options.Debug)
                        Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Зображення успішно отримані. Спроба нанесення титрів...");
                    for (int j = 0; j < exportdata.GetLength(0); j++)
                    {
                        try
                        {
                            using (Image inputImage = Image.FromFile(Path.GetTempPath() + id + exportdata[j, 0] + ".bmp"))
                            {
                                #region Image Subtitles
                                using (System.Drawing.Graphics graphicsInputImage = Graphics.FromImage(inputImage))
                                {

                                    System.Drawing.Drawing2D.GraphicsPath myPath = new System.Drawing.Drawing2D.GraphicsPath();

                                    // Set up all the string parameters.
                                    string stringText = titles.Replace(@"\r\n", Environment.NewLine);
                                    FontFamily family = new FontFamily(Options.Font);
                                    int fontStyle = (int)FontStyle.Regular;
                                    int emSize = Options.FontSize;
                                    Point origin = new Point(10, 10);
                                    StringFormat format = StringFormat.GenericDefault;

                                    // Add the string to the path.
                                    myPath.AddString(stringText,
                                        family,
                                        fontStyle,
                                        emSize,
                                        origin,
                                        format);


                                    SizeF size = TextRenderer.MeasureText("Камера №" + exportdata[j, 0] + " - " + camerainfo[j], new Font(new FontFamily(Options.Font), Options.FontSize, FontStyle.Regular, GraphicsUnit.Point));
                                    origin = new Point(inputImage.Width - (int)size.Width, inputImage.Height - (int)size.Height);
                                    myPath.AddString("Камера №" + exportdata[j, 0] + " - " + camerainfo[j], family, fontStyle, emSize, origin, format);

                                    graphicsInputImage.DrawString("Розроблено aydnep@aydnep.com.ua©", new Font("Arial", 8, FontStyle.Regular), new SolidBrush(Color.FromArgb(100, 255, 255, 255)), new PointF(20.0F, inputImage.Height - 15.0F), new StringFormat());
                                    graphicsInputImage.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                                    graphicsInputImage.DrawPath(new Pen(new SolidBrush(Options.TextColor), Options.StrokeSize), myPath);
                                    graphicsInputImage.FillPath(new SolidBrush(Options.StrokeColor), myPath);

                                    //Bitmap ap = new Bitmap(AP.IntegrationTools.Properties.Resources.ap);


                                    if (Options.Debug)
                                        Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Нанесення титрів " + titles);

                                } // using image for titles
                                #endregion
                                #region Image Save with subtitles
                                switch (Options.Imgformat)
                                {
                                    case "bmp":
                                        inputImage.Save(exportdata[j, 1] + ".bmp", ImageFormat.Bmp);
                                        if (Options.Debug)
                                            Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Збереження зображення у файл: " + exportdata[j, 1] + ".bmp");
                                        break;
                                    case "gif":
                                        inputImage.Save(exportdata[j, 1] + ".gif", ImageFormat.Gif);
                                        if (Options.Debug)
                                            Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Збереження зображення у файл: " + exportdata[j, 1] + ".gif");
                                        break;
                                    case "jpg":
                                        inputImage.Save(exportdata[j, 1] + ".jpg", ImageFormat.Jpeg);
                                        if (Options.Debug)
                                            Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Збереження зображення у файл: " + exportdata[j, 1] + ".jpg");
                                        break;
                                    case "png":
                                        inputImage.Save(exportdata[j, 1] + ".png", ImageFormat.Png);
                                        if (Options.Debug)
                                            Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Збереження зображення у файл: " + exportdata[j, 1] + ".png");
                                        break;
                                    case "tiff":
                                        inputImage.Save(exportdata[j, 1] + ".tiff", ImageFormat.Tiff);
                                        if (Options.Debug)
                                            Log.Write("AP.CCTV.iidk", "CheckFile", "Збереження зображення у файл: " + exportdata[j, 1] + ".tiff");
                                        break;
                                    default:
                                        inputImage.Save(exportdata[j, 1] + ".jpg", ImageFormat.Jpeg);
                                        if (Options.Debug)
                                            Log.Write("AP.CCTV.iidk", "CheckFile", "Збереження зображення у файл: " + exportdata[j, 1] + ".jpg");

                                        break;
                                } // switch
                                #endregion
                            } //using tmpimage
                            if (Options.Debug)
                                Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Спроба видалити тимчасовий файл " + id + exportdata[j, 0] + ".bmp");
                            if (File.Exists(Path.GetTempPath() + id + exportdata[j, 0] + ".bmp"))
                                try
                                {
                                    File.Delete(Path.GetTempPath() + id + exportdata[j, 0] + ".bmp");
                                }
                                catch (Exception ex)
                                {
                                    Log.Write("AP.IntegrationTools.iidk", "CheckFile", ex.Message);
                                }

                        }
                        catch (Exception ex)
                        {
                            Log.Write("AP.IntegrationTools.iidk", "CheckFile", ex.Message);
                            are.Set();
                        }
                        if (Options.Debug)
                            Log.Write("AP.IntegrationTools.iidk", "CheckFile", "Оброблено зображень: " + (j + 1) + " з " + exportdata.GetLength(0));
                    } // for (j 0 to requested files) titles on image
                    result = true;
                    are.Set();
                } // if (tmp files = requested files)
                #endregion
            } // for (i 0 to 30)
            are.Set();
        }
        #endregion

        #region COM+ visible methods
        /// <summary>
        /// Виклик допомоги, щодо доступних методів у інтерфейсі
        /// </summary>
        /// <returns>Повертає строку зі списком методів, доступних у данному інтерфейсі</returns>
        public string Help()
        {
            MethodInfo[] methodInfos = this.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
            StringBuilder methods = new StringBuilder();
            methods.AppendLine(AssemblyAbout.GetVersion());
            foreach (MethodInfo method in methodInfos)
            {
                methods.Append(method.ReturnType + " " + method.Name + "(");
                if (method.GetParameters().Length > 0)
                {
                    foreach (ParameterInfo parameter in method.GetParameters())
                    {
                        methods.Append(parameter.ParameterType + " " + parameter.Name + ",");
                    }
                    methods.Remove(methods.Length - 1, 1);
                }
                methods.AppendLine(")");
            }
            return methods.ToString();
        }
        /// <summary>
        /// Функція ініціалізації компоненти інтеграції систем відеонагляду та 1Сv8
        /// </summary>
        /// <param name="ip">IP-адреса сервеа відеонагляду (string)</param>
        /// <param name="cameras">Двомірний масив строк {№камери, повний шлях до файлу без розширення} (string[,])</param>
        /// <param name="title">Субтитри, що накладаються на зображення (для переносу строк викристовуться символи "\r\n") (string)</param>
        /// <param name="weightTime">Період часу з якого потрібно отримати зображення (якщо пустий то в реальному часі)</param>
        /// <value>Повертає істину або хибність отримання зображення</value>
        public bool SaveImage(string ip, string[,] cameras, string title, string weightTime)
        {
            Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Init");
#if (DEBUG)
            {
                Options.Debug = true;
            }
#else
            { }
#endif  
            id = "thread" + Guid.NewGuid().ToString();
            cctvIP = ip;
            exportdata = cameras;
            titles = title;
            exportTime = weightTime;
            result = false;
            Log.Write("AP.IntegrationTools.iidk", "SaveImage", "IP: " + ip + ", Titles: " + title + ", weightTime: " + weightTime);
            try
            {
                for (int i = 0; i < exportdata.GetLength(0); i++)
                {
                    if (File.Exists(exportdata[i, 1] + "." + Options.Imgformat))
                        File.Delete(exportdata[i, 1] + "." + Options.Imgformat);
                    if (File.Exists(Path.GetTempPath() + id + exportdata[i, 0] + ".bmp"))
                        File.Delete(Path.GetTempPath() + id + exportdata[i, 0] + ".bmp");
                }
            }
            catch (Exception ex)
            {
                Log.Write("AP.IntegrationTools.iidk", "SaveImage", ex.Message);
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Помилка! Заблоковані тимчасові файли!");
            }
            if (!CheckParams())
            {
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Отримана помилка при перевірці параметрів. Подальша робото неможлива.");
                return false;
            }
            camerainfo = new string[exportdata.GetLength(0)];
            WaitHandle[] wh = new WaitHandle[]
            {
                new AutoResetEvent(false)
            };
            ThreadPool.QueueUserWorkItem(new WaitCallback(CheckFile), wh[0]);
            if (Options.Debug)
                Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Ініціалізація компоненти...");
            try
            {
                ocx = new AxIIDK_COMLib.AxIIDK_COM();
                ocx.CreateControl();
                ocx.OnConnect += ocx_OnConnect;
                ocx.OnVideoFrame += ocx_OnVideoFrame;

                ocx.Options = 1;
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Спроба підключення до " + cctvIP + "...");
                ocx.AsyncConnectUnique(cctvIP, 900);
                WaitHandle.WaitAll(wh);
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Знищення компоненети...");
                ocx.Disconnect();
                ocx.Dispose();
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Компонента знищена.");
            }
            catch (Exception ex)
            {
                Log.Write("AP.IntegrationTools.iidk", "SaveImage", ex.Message);
            }
            try
            {
                for (int i = 0; i < exportdata.GetLength(0); i++)
                {
                    if (File.Exists(Path.GetTempPath() + id + exportdata[i, 0] + ".bmp"))
                        File.Delete(Path.GetTempPath() + id + exportdata[i, 0] + ".bmp");
                }
            }
            catch (Exception ex)
            {
                Log.Write("AP.IntegrationTools.iidk", "SaveImage", ex.Message);
                if (Options.Debug)
                    Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Помилка! Заблоковані тимчасові файли!");
            }
            if (Options.Debug)
                Log.Write("AP.IntegrationTools.iidk", "SaveImage", "Повернення результату " + result);
            return result;
        }
        /// <summary>
        /// Функція для зміни параметрів за замовченням
        /// В залежності від якості каналу зв’язку рекомендовані наступні параметри:
        /// &gt;10Mbs - кількість спроб: 20; затримка: 500; стиснення - 0.
        /// &gt;2Mbs &lt;10Mbs - кількість спроб: 30; затримка: 1000; стиснення - 1.
        /// &lt;2Mbs - кількість спроб: 120; затримка: 3000; стиснення - 1..5.
        /// </summary>
        /// <param name="count">Кількість спроб (int32)</param>
        /// <param name="compression">Розмір компрессії (0 - без зтиснення, 5 - максимальне стиснення) (int32)</param>
        /// <param name="imgformat">Формат збереження зображень (jpg, gif, png, tiff, bmp) (string)</param>
        /// <param name="debug">Режим відладки</param>
        public void SetExportParams(int count, int compression, string imgformat, bool debug)
        {
            Options.Count = count;
            Options.Compression = compression;
            Options.Imgformat = imgformat.ToLower();
            Options.Debug = debug;
        }
        /// <summary>
        /// Функція для зміни облікових данних за замовченням
        /// </summary>
        /// <param name="login">Користувач (string)</param>
        /// <param name="password">Пароль (string)</param>
        public void SetCerdentials(string login, string password)
        {
            Options.Login = login;
            Options.Password = password;
        }
        /// <summary>
        /// Функція для маркування легальності проїзду транспорту
        /// </summary>
        /// <param name="ip">IP-адреса серверу відеоспостереження</param>
        /// <param name="database">Назва бази данних</param>
        /// <param name="id">Номер ТТН</param>
        /// <param name="carID">Номер автомобіля</param>
        /// <param name="entry">Номер проїзду</param>
        /// <param name="culture">Культура</param>
        /// <param name="arm">Етап бізнес-процесу</param>
        public void SendLegal(string ip, string database, string id, string carID, int entry, string culture, string arm)
        {
            SqlConnection db = new SqlConnection();
            db.ConnectionString = "Data Source=" + ip + ";Initial Catalog=" + database + ";" + "user id=" + Options.Login + ";password=" + Options.Password + ";";
            try
            {
                db.Open();
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Could not connect to Database");
                Log.Write("AP.IntegrationTools.iidk", "SendLegal", "Could not connect to Database");
                Log.Write("AP.IntegrationTools.iidk", ex.Source, ex.Message);
                Log.Write("AP.IntegrationTools.iidk", "SqlConnection", db.ConnectionString);
                return;
            }
            SqlCommand sql = new SqlCommand("INSERT INTO [PROTOCOL1C] ([ID],[TruckID],[Entry],[Culture],[WHO]) VALUES", db);
            sql.CommandText += "('" + id + "','" + carID + "'," + entry + ",'" + culture + "','" + arm + "')";
            try
            {
                sql.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Log.Write("AP.IntegrationTools.iidk", "SendLegal", "SQL syntax error. Please Check");
                Log.Write("AP.IntegrationTools.iidk", ex.Source, ex.Message);
                Log.Write("AP.IntegrationTools.iidk", "SqlCommand", sql.CommandText);
                return;
            }
        }
        /// <summary>
        /// Функція для зміни параметрів титрів
        /// </summary>
        /// <param name="font">Назва шрифту</param>
        /// <param name="size">Розмір шрифту</param>
        /// <param name="textcolor">Кольор шрифту</param>
        /// <param name="strokecolor">Кольор обведення</param>
        /// <param name="strokesize">Розмір обведення</param>
        public void SetTitleOptions(string font, int size, string textcolor, string strokecolor, int strokesize)
        {
            Options.Font = font;
            Options.FontSize = size;
            Options.TextColor = Color.FromName(textcolor);
            Options.StrokeColor = Color.FromName(strokecolor);
            Options.StrokeSize = strokesize;
        }
        /// <summary>
        /// Функція повертає список системних кольорів
        /// </summary>
        /// <returns>string[] масив назв кольорів</returns>
        public string[] GetColors()
        {
            KnownColor[] colors = (KnownColor[])Enum.GetValues(typeof(KnownColor));
            List<string> result = new List<string>();
            foreach (KnownColor knowColor in colors)
            {
                result.Add(knowColor.ToString());
            }
            return result.ToArray();
        }

        #endregion
    }

    #endregion

    #region Internal Helpers

    static class Options
    {
        #region Safe variables
        public static int Count
        {
            get { return count; }
            set { count = value; }
        }
        public static int Compression
        {
            get { return compression; }
            set { compression = value; }
        }
        public static string Imgformat
        {
            get { return imgformat; }
            set { imgformat = value; }
        }
        public static bool Debug
        {
            get { return debug; }
            set { debug = value; }
        }
        public static string Login
        {
            get { return login; }
            set { login = value; }
        }
        public static string Password
        {
            get { return password; }
            set { password = value; }
        }
        public static string Font
        {
            get { return font; }
            set { font = value; }
        }
        public static int FontSize
        {
            get { return fontsize; }
            set { fontsize = value; }
        }
        public static Color TextColor
        {
            get { return textcolor; }
            set { textcolor = value; }
        }
        public static Color StrokeColor
        {
            get { return strokecolor; }
            set { strokecolor = value; }
        }
        public static int StrokeSize
        {
            get { return strokesize; }
            set { strokesize = value; }
        }
        #endregion
        #region private variables
        static int count = 30;
        static int compression = 1;
        static string imgformat = "jpg";
        static bool debug = false;
        static string login = "video";
        static string password = "P@ssw0rd1";
        static string font = "Arial";
        static int fontsize = 12;
        static Color textcolor = Color.Black;
        static Color strokecolor = Color.Yellow;
        static int strokesize = 2;
        #endregion
    }

    class Log
    {
        private static object sync = new object();
        public static void Write(string exDeclaringType, string exName, string exMessage)
        {
            try
            {
                // Путь .\\Log
                string pathToLog = Path.Combine(Path.GetTempPath(), "AP.IntegrationTools");
                if (!Directory.Exists(pathToLog))
                    Directory.CreateDirectory(pathToLog); // Создаем директорию, если нужно
                string filename = Path.Combine(pathToLog, string.Format("{0}_{1:dd.MM.yyy}.log",
                "iidk", DateTime.Now));
                string fullText = string.Format("[{0:dd.MM.yyy HH:mm:ss.fff}] [{1}.{2}()] {3}\r\n",
                DateTime.Now, exDeclaringType, exName, exMessage);
                lock (sync)
                {
                    File.AppendAllText(filename, fullText, Encoding.GetEncoding("Windows-1251"));
                }
            }
            catch
            {
                // Перехватываем все и ничего не делаем
            }
        }
    }

    [StructLayout(LayoutKind.Sequential), ComVisible(false)]
    internal class SystemTime
    {
        public ushort wYear;
        public ushort wMonth;
        public ushort wDayOfWeek;
        public ushort wDay;
        public ushort wHour;
        public ushort wMinute;
        public ushort wSecond;
        public ushort wMilliseconds;
        public DateTime ToDateTime()
        {
            return new DateTime(wYear, wMonth, wDay, wHour, wMinute, wSecond, wMilliseconds);
        }
    }

    internal class AssemblyAbout
    {
        internal static string GetVersion()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
            return (fvi.FileVersion);
        }

        internal static string GetPublicMethodsOfType(string type)
        {
            MethodInfo[] methodInfos = Type.GetType(type).GetMethods(BindingFlags.Public | BindingFlags.Instance);
            StringBuilder methods = new StringBuilder();
            foreach (MethodInfo method in methodInfos)
            {
                methods.AppendLine(method.Name);
            }
            return methods.ToString();
        }
    }
    #endregion
}
