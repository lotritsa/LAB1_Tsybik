using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExel
{
    public partial class Form2 : Form
    {
        Form1 form;
        public Form2(Form1 f)
        {
            InitializeComponent();
            this.form = f;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = @"C:\exel_saves\";
            string pathEnding = ".txt";
           string name = textBox2.Text.ToString();
            if (File.Exists(path + name))
            {
                path = path + name;
            } else
            {
                if (File.Exists(path+name+pathEnding))
                {
                    path = path + name + pathEnding;
                } else
                {
                    string message = "File not found!";
                    string caption = "Alert";
                    MessageBoxButtons buttons = MessageBoxButtons.OK;
                    MessageBox.Show(message, caption, buttons);
                    this.Close();
                    return;
                }
            }
            string text =  "";
            using (StreamReader sr = new StreamReader(path, System.Text.Encoding.Default))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    text = text + line + '\n';
                }
            }
            this.form.TableInitialize(100,100);
            this.form.StringToTable(text);
            this.Close();
        }
    }
}
