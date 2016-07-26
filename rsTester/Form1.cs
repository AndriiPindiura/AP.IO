using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;

namespace rsTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            z2usb = new AP.IntegrationTools.Rs232();
            z2usb.OnDataRead += z2DataRecieved;

            /*try
            {
                iidkCOM = Activator.CreateInstance(Type.GetTypeFromProgID("AP.CCTV.iidk"));
                richTextBox1.AppendText(Type.GetTypeFromProgID("AP.CCTV.iidk").Assembly.ToString() + Environment.NewLine + Environment.NewLine);
                //iidkCOM.OnDataRecieved += z2DataRecieved;
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText("Помилка при підключенні до COM застосунку:" + Environment.NewLine);
                richTextBox1.AppendText(ex.Message + Environment.NewLine);
                MessageBox.Show(ex.Message);
            }*/
        }

        private dynamic iidkCOM;

        private AP.IntegrationTools.Rs232 z2usb;

        private void z2DataRecieved(string msg)
        {
            richTextBox1.Invoke((MethodInvoker)delegate
            {
                richTextBox1.AppendText("Зчитана карта №" + msg);
            });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AP.IntegrationTools.Rs232 rs = new AP.IntegrationTools.Rs232();
            if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Необхідно обрати порт та тип вагопроцесору!");
                comboBox1.Focus();
                return;
            }
            richTextBox1.AppendText("Спроба отримати вагу з: " + comboBox2.SelectedItem.ToString() + " на порту: " + comboBox1.SelectedItem.ToString() + "\r\n");
            try
            {
                richTextBox1.AppendText("Отримана вага: " + rs.GetWeight(comboBox1.SelectedItem.ToString(), comboBox2.SelectedIndex + 1, Convert.ToInt32(numericUpDown1.Value)).ToString() + " кг\r\n");
            }
            catch (Exception ex)
            {
                richTextBox1.AppendText("Не вдалось отримаги вагу!\r\n");
                if (ex.InnerException != null)
                    richTextBox1.AppendText(ex.InnerException.Message + "\r\n");
                else
                    richTextBox1.AppendText(ex.Message + "\r\n");

            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.ScrollToCaret();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.AddRange(SerialPort.GetPortNames());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Необхідно обрати порт!");
                comboBox1.Focus();
                return;
            }
            try
            {
                z2usb.OpenZ2usb(comboBox1.SelectedItem.ToString());
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                    richTextBox1.AppendText(ex.InnerException.Message + "\r\n");
                else
                    richTextBox1.AppendText(ex.Message + "\r\n");

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            DateTime start = DateTime.Now;
            AP.IntegrationTools.Infratec infratec = new AP.IntegrationTools.Infratec();
            string analyze = infratec.GetGrainAnalyze(textBox1.Text, (int)numericUpDown3.Value, 30, 5);
            richTextBox1.Text = analyze;
            richTextBox1.AppendText((DateTime.Now - start).ToString() + "\r\n");
            string[] results = infratec.ParseGrainAnalyze(analyze);
            for (int count = 0; count < results.GetLength(0); count++)
            {
                richTextBox1.AppendText(results[count] + "\r\n");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            DateTime start = DateTime.Now;
            AP.IntegrationTools.MoxaIO moxa = new AP.IntegrationTools.MoxaIO();
            try
            {
                bool[] status = moxa.GetDIStatus(textBox1.Text, (ushort)numericUpDown3.Value, 5000, "", (byte)numericUpDown2.Value);
                for (int i = 0; i < status.Length; i++)
                {
                    richTextBox1.AppendText(string.Format("IO({0}) status: {1}\r\n", i, status[i] ? "ON" : "OFF"));
                }
            }
            catch (Exception ex)
            {
                if (ex.InnerException != null)
                    MessageBox.Show(ex.InnerException.Message);
                else
                    MessageBox.Show(ex.Message);
            }

        }
    }
}
