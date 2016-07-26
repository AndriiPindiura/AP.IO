using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MOXA_CSharp_MXIO;

namespace MoxaIO   
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public const UInt16 Port = 502;						//Modbus TCP port
        public const UInt16 DO_SAFE_MODE_VALUE_OFF = 0;
        public const UInt16 DO_SAFE_MODE_VALUE_ON = 1;
        public const UInt16 DO_SAFE_MODE_VALUE_HOLD_LAST = 2;

        public const UInt16 DI_DIRECTION_DI_MODE = 0;
        public const UInt16 DI_DIRECTION_COUNT_MODE = 1;
        public const UInt16 DO_DIRECTION_DO_MODE = 0;
        public const UInt16 DO_DIRECTION_PULSE_MODE = 1;

        public const UInt16 TRIGGER_TYPE_LO_2_HI = 0;
        public const UInt16 TRIGGER_TYPE_HI_2_LO = 1;
        public const UInt16 TRIGGER_TYPE_BOTH = 2;
        //A-OPC Server response W5340 Device STATUS information data filed index
        public const int IP_INDEX = 0;
        public const int MAC_INDEX = 4;

        private static string CheckErr(int iRet, string szFunctionName)
        {
            string szErrMsg = "MXIO_OK";

            if (iRet != MXIO_CS.MXIO_OK)
            {

                switch (iRet)
                {
                    case MXIO_CS.ILLEGAL_FUNCTION:
                        szErrMsg = "ILLEGAL_FUNCTION";
                        break;
                    case MXIO_CS.ILLEGAL_DATA_ADDRESS:
                        szErrMsg = "ILLEGAL_DATA_ADDRESS";
                        break;
                    case MXIO_CS.ILLEGAL_DATA_VALUE:
                        szErrMsg = "ILLEGAL_DATA_VALUE";
                        break;
                    case MXIO_CS.SLAVE_DEVICE_FAILURE:
                        szErrMsg = "SLAVE_DEVICE_FAILURE";
                        break;
                    case MXIO_CS.SLAVE_DEVICE_BUSY:
                        szErrMsg = "SLAVE_DEVICE_BUSY";
                        break;
                    case MXIO_CS.EIO_TIME_OUT:
                        szErrMsg = "EIO_TIME_OUT";
                        break;
                    case MXIO_CS.EIO_INIT_SOCKETS_FAIL:
                        szErrMsg = "EIO_INIT_SOCKETS_FAIL";
                        break;
                    case MXIO_CS.EIO_CREATING_SOCKET_ERROR:
                        szErrMsg = "EIO_CREATING_SOCKET_ERROR";
                        break;
                    case MXIO_CS.EIO_RESPONSE_BAD:
                        szErrMsg = "EIO_RESPONSE_BAD";
                        break;
                    case MXIO_CS.EIO_SOCKET_DISCONNECT:
                        szErrMsg = "EIO_SOCKET_DISCONNECT";
                        break;
                    case MXIO_CS.PROTOCOL_TYPE_ERROR:
                        szErrMsg = "PROTOCOL_TYPE_ERROR";
                        break;
                    case MXIO_CS.SIO_OPEN_FAIL:
                        szErrMsg = "SIO_OPEN_FAIL";
                        break;
                    case MXIO_CS.SIO_TIME_OUT:
                        szErrMsg = "SIO_TIME_OUT";
                        break;
                    case MXIO_CS.SIO_CLOSE_FAIL:
                        szErrMsg = "SIO_CLOSE_FAIL";
                        break;
                    case MXIO_CS.SIO_PURGE_COMM_FAIL:
                        szErrMsg = "SIO_PURGE_COMM_FAIL";
                        break;
                    case MXIO_CS.SIO_FLUSH_FILE_BUFFERS_FAIL:
                        szErrMsg = "SIO_FLUSH_FILE_BUFFERS_FAIL";
                        break;
                    case MXIO_CS.SIO_GET_COMM_STATE_FAIL:
                        szErrMsg = "SIO_GET_COMM_STATE_FAIL";
                        break;
                    case MXIO_CS.SIO_SET_COMM_STATE_FAIL:
                        szErrMsg = "SIO_SET_COMM_STATE_FAIL";
                        break;
                    case MXIO_CS.SIO_SETUP_COMM_FAIL:
                        szErrMsg = "SIO_SETUP_COMM_FAIL";
                        break;
                    case MXIO_CS.SIO_SET_COMM_TIME_OUT_FAIL:
                        szErrMsg = "SIO_SET_COMM_TIME_OUT_FAIL";
                        break;
                    case MXIO_CS.SIO_CLEAR_COMM_FAIL:
                        szErrMsg = "SIO_CLEAR_COMM_FAIL";
                        break;
                    case MXIO_CS.SIO_RESPONSE_BAD:
                        szErrMsg = "SIO_RESPONSE_BAD";
                        break;
                    case MXIO_CS.SIO_TRANSMISSION_MODE_ERROR:
                        szErrMsg = "SIO_TRANSMISSION_MODE_ERROR";
                        break;
                    case MXIO_CS.PRODUCT_NOT_SUPPORT:
                        szErrMsg = "PRODUCT_NOT_SUPPORT";
                        break;
                    case MXIO_CS.HANDLE_ERROR:
                        szErrMsg = "HANDLE_ERROR";
                        break;
                    case MXIO_CS.SLOT_OUT_OF_RANGE:
                        szErrMsg = "SLOT_OUT_OF_RANGE";
                        break;
                    case MXIO_CS.CHANNEL_OUT_OF_RANGE:
                        szErrMsg = "CHANNEL_OUT_OF_RANGE";
                        break;
                    case MXIO_CS.COIL_TYPE_ERROR:
                        szErrMsg = "COIL_TYPE_ERROR";
                        break;
                    case MXIO_CS.REGISTER_TYPE_ERROR:
                        szErrMsg = "REGISTER_TYPE_ERROR";
                        break;
                    case MXIO_CS.FUNCTION_NOT_SUPPORT:
                        szErrMsg = "FUNCTION_NOT_SUPPORT";
                        break;
                    case MXIO_CS.OUTPUT_VALUE_OUT_OF_RANGE:
                        szErrMsg = "OUTPUT_VALUE_OUT_OF_RANGE";
                        break;
                    case MXIO_CS.INPUT_VALUE_OUT_OF_RANGE:
                        szErrMsg = "INPUT_VALUE_OUT_OF_RANGE";
                        break;
                }

                Console.WriteLine("Function \"{0}\" execution Fail. Error Message : {1}\n", szFunctionName, szErrMsg);

                if (iRet == MXIO_CS.EIO_TIME_OUT || iRet == MXIO_CS.HANDLE_ERROR)
                {
                    //To terminates use of the socket
                    MXIO_CS.MXEIO_Exit();
                }
            }
            return szErrMsg;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            int ret = MXIO_CS.MXIO_GetDllVersion();
            int major = (ret >> 12) & 0xF;
            int minor = (ret >> 8) & 0xF;
            int build = (ret >> 4) & 0xF;
            int revision = (ret) & 0xF;
            Int32[] hConnection = new Int32[1];
            string version = String.Format("{0}.{1}.{2}.{3}", major, minor, build, revision);
            richTextBox1.AppendText(version + Environment.NewLine);
            ret = MXIO_CS.MXEIO_Init();
            richTextBox1.AppendText(ret.ToString() + Environment.NewLine);
            MXIO_CS.MXEIO_E1K_Connect(UTF8Encoding.UTF8.GetBytes(textBox1.Text), 502, 5000, hConnection, UTF8Encoding.UTF8.GetBytes(""));
            if (ret == MXIO_CS.MXIO_OK)
                richTextBox1.AppendText("MXEIO_E1K_Connect Success.\r\n");
            byte[] bytCheckStatus = new byte[1];
            ret = MXIO_CS.MXEIO_CheckConnection(hConnection[0], 5000, bytCheckStatus);
            //CheckErr(ret, "MXEIO_CheckConnection");
            if (ret == MXIO_CS.MXIO_OK)
            {
                switch (bytCheckStatus[0])
                {
                    case MXIO_CS.CHECK_CONNECTION_OK:
                        richTextBox1.AppendText(String.Format("MXEIO_CheckConnection: Check connection ok => {0}\r\n", bytCheckStatus[0]));
                        break;
                    case MXIO_CS.CHECK_CONNECTION_FAIL:
                        richTextBox1.AppendText(String.Format("MXEIO_CheckConnection: Check connection fail => {0}\r\n", bytCheckStatus[0]));
                        break;
                    case MXIO_CS.CHECK_CONNECTION_TIME_OUT:
                        richTextBox1.AppendText(String.Format("MXEIO_CheckConnection: Check connection time out => {0}\r\n", bytCheckStatus[0]));
                        break;
                    default:
                        richTextBox1.AppendText(String.Format("MXEIO_CheckConnection: Check connection status unknown => {0}\r\n", bytCheckStatus[0]));
                        break;
                }
            }
            byte bytCount = 4;
            byte bytStartChannel = 0;
            UInt32[] dwGetDIValue = new UInt32[1];
            ret = MXIO_CS.E1K_DI_Reads(hConnection[0], bytStartChannel, bytCount, dwGetDIValue);
            if (ret == MXIO_CS.MXIO_OK)
            {
                //Console.WriteLine("E1K_DI_Reads Get Ch0~ch3 DI Direction DI Mode DI Value success.");
                for (int i = 0, dwShiftValue = 0; i < bytCount; i++, dwShiftValue++)
                    richTextBox1.AppendText(String.Format("DI vlaue: ch[{0}] = {1}", i + bytStartChannel, ((dwGetDIValue[0] & (1 << dwShiftValue)) == 0) ? "OFF" : "ON") + Environment.NewLine);
            }
            //--------------------------------------------------------------------------
            //End Application
            ret = MXIO_CS.MXEIO_Disconnect(hConnection[0]);
            if (ret == MXIO_CS.MXIO_OK)
                richTextBox1.AppendText(String.Format("MXEIO_Disconnect return {0}\r\n", ret));
            //--------------------------------------------------------------------------
            MXIO_CS.MXEIO_Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}
