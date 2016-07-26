using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;

namespace COMregister
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string framework = Environment.GetEnvironmentVariable("SystemRoot") + @"\Microsoft.NET\Framework\v4.0.30319\";


        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                button2.Enabled = true;
                button3.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo(textBox1.Text);
            RegistrationServices regAsm = new RegistrationServices();
            /*Process p = new Process();
            p.StartInfo.FileName = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory() + "regasm.exe";
            p.StartInfo.Arguments = fi.FullName + " /codebase";
            p.Start();
            p.WaitForExit();
            if (p.ExitCode == 0)
                MessageBox.Show("Вдала реестрація типів!");
            else
                MessageBox.Show("Помилка при реестрації типів: " + p.ExitCode.ToString());*/
            try
            {
                bool result = regAsm.RegisterAssembly(Assembly.LoadFile(fi.FullName), AssemblyRegistrationFlags.SetCodeBase);
                if (result)
                {
                    MessageBox.Show("Вдала реестрація типів!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            FileInfo fi = new FileInfo(textBox1.Text);
            Process p = new Process();
            p.StartInfo.FileName = System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory() + "regasm.exe";
            p.StartInfo.Arguments = fi.FullName + " /unregister";
            p.Start();
            p.WaitForExit();
            if (p.ExitCode == 0)
                MessageBox.Show("Вдала реестрація типів!");
            else
                MessageBox.Show("Помилка при реестрації типів: " + p.ExitCode.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                dynamic comObject = Activator.CreateInstance(Type.GetTypeFromProgID(textBox1.Text));
                MessageBox.Show(Type.GetTypeFromProgID(textBox1.Text).Assembly.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
