using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.IO;

namespace AP.CCTV.Debug
{
    public partial class iidkDebug : Form
    {
        public iidkDebug()
        {
            InitializeComponent();
            cctvForm = new AP.IntegrationTools.iidk();
        }

        private AP.IntegrationTools.iidk cctvForm;
        private DateTime total;

        private void ImageRecieved(bool image)
        {
            richTextBox1.Invoke((MethodInvoker)delegate
            {
                if (image)
                {
                    //MessageBox.Show("GetImage");
                    //cctvForm.Close();
                    richTextBox1.AppendText("Зображення успішно отримані.\r\nФайли сформовані. Затрачено часу: " + (DateTime.Now - total).ToString() + "\r\n");
                }
                else
                {
                    //MessageBox.Show("Error");
                    //cctvForm.Close();
                    richTextBox1.AppendText("Не вдалось отримати зобораження.\r\nКількість спроб: " + ((Int32)numericUpDown1.Value).ToString() + " . Затрачено часу: " + (DateTime.Now - total).ToString() + "\r\n");
                }
            });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            total = DateTime.Now;
            for (int j = 1; j <= numericUpDown2.Value; j++)
            {
                richTextBox1.Clear();
                richTextBox1.AppendText("Спроба отримати зображення з " + textBox1.Text + " Попытка №" + j + "\r\n");
                string[] cams = textBox2.Text.Split(';');
                string[,] datatoexport = new string[cams.Length, cams.Length];
                for (int i = 0; i < cams.Length; i++)
                {
                    datatoexport[i, 0] = cams[i];
                    datatoexport[i, 1] = Path.GetTempPath() + Guid.NewGuid().ToString();
                }
                //MessageBox.Show(datatoexport.GetLength(0) + "\r\n" + datatoexport.GetLength(1));
                if (cctvForm.SaveImage(textBox1.Text, datatoexport, textBox3.Text, textBox5.Text))
                    richTextBox1.AppendText("Зображення успішно отримані.\r\nФайли сформовані. Затрачено часу: " + (DateTime.Now - total).ToString() + "\r\n");
                else
                    richTextBox1.AppendText("Не вдалось отримати зобораження.\r\nКількість спроб: " + ((Int32)numericUpDown1.Value).ToString() + " . Затрачено часу: " + (DateTime.Now - total).ToString() + "\r\n");
                //cctvForm.Destroy();
            }
            richTextBox1.AppendText("Затрачено часу: " + (DateTime.Now - total).ToString() + "\r\n");

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.ScrollToCaret();
        }

        private void cctvDebug_Load(object sender, EventArgs e)
        {
            richTextBox1.AppendText("Поточне місце для збереження файлів: " + Path.GetTempPath() + "\r\n");
            foreach (string s in cctvForm.GetColors())
            {
                comboBox2.Items.Add(s);
                comboBox3.Items.Add(s);
            }
            comboBox1.SelectedIndex = 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cctvForm.SetExportParams((Int32)numericUpDown1.Value, (Int32)numericUpDown3.Value, comboBox1.SelectedItem.ToString(), true);
        }


        private void button3_Click_1(object sender, EventArgs e)
        {
            cctvForm.SetTitleOptions(textBox4.Text, (Int32)numericUpDown4.Value, comboBox2.SelectedItem.ToString(), comboBox3.SelectedItem.ToString(), (Int32)numericUpDown5.Value);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (File.Exists(saveFileDialog1.FileName))
                File.Delete(saveFileDialog1.FileName);
            foreach (string s in cctvForm.GetColors())
            {
                File.AppendAllText(saveFileDialog1.FileName, s + Environment.NewLine);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            cctvForm.SendLegal(textBox1.Text + "\\SQLEXPRESS", "Intellect", "ttn", "car", 1, "culture", "test");
        }

    }
}
