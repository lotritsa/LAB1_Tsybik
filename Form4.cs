using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Collections;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExel
{
    public partial class Form4 : Form
    {
        MyTable myHash;
        public Form4(MyTable mh)
        {
            InitializeComponent();
            this.myHash = mh;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string enterText = textBox2.Text;
            int numSymbols = 0;
            int wrongSym = 0;
            for (int i = 0; i<enterText.Length; i++)
            {
                if (enterText[i] >= 'A' && enterText[i]<='Z')
                {
                    numSymbols++;
                } else
                {
                    if (enterText[i] >= 'a' && enterText[i] <= 'z')
                    {
                        numSymbols++;
                    } else
                    {
                        if (enterText[i] >= '0' && enterText[i]<='9')
                        {
                            numSymbols++;
                        } 
                        else
                        {
                            wrongSym++;
                        }
                    }
                }
            }
            if (numSymbols == 0 || wrongSym>0)
            {
                string message = "Invalid file name!";
                string caption = "Alert";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                MessageBox.Show(message, caption, buttons);

            } else
            {

                string fileName = @"C:\exel_saves\";
                string nameEnding = ".txt";
                fileName = fileName + enterText + nameEnding;
                if (File.Exists(fileName))
                {
                    string message = "This file already exists. Overwrite the file?";
                    string caption = "Alert";
                    MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                    if (MessageBox.Show(message, caption, buttons) == System.Windows.Forms.DialogResult.No)
                    {
                        return;
                    }
                }
                File.Delete(fileName);
                string filetext = "";
                foreach (DictionaryEntry pair in myHash.Expressions)
                {

                    string index = pair.Key.ToString();
                    filetext = filetext + index + ' ' + pair.Value + "\n";
                }
                if (filetext.Length > 4)
                {
                    using (StreamWriter sw = File.CreateText(fileName))
                    {
                        sw.WriteLine(filetext);
                    }
                    string message1 = "File has been successfully saved!";
                    string caption1 = "Alert";
                    MessageBoxButtons buttons1 = MessageBoxButtons.OK;
                    MessageBox.Show(message1, caption1, buttons1);
                    this.Close();

                }
                else
                {
                    string message1 = "Saving failed!";
                    string caption1 = "Alert";
                    MessageBoxButtons buttons1 = MessageBoxButtons.OK;
                    MessageBox.Show(message1, caption1, buttons1);
                    this.Close();
                }
            }
                
        }
    }
}
