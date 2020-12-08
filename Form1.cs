using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExel
{
    public partial class Form1 : Form
    {
        MyTable myHash;
        Parser converter;
        Base10fromto26 f;
        public Form1()
        {
            myHash = MyTable.GetInstance();
            converter = new Parser();
            f = new Base10fromto26();
            InitializeComponent();
            TableInitialize(50,15);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        public void TableInitialize(int columnsNum, int rowsNum)
        {
            textBox1.Text = "hi";
            dataGridView1.Columns.Clear();
            dataGridView1.ColumnHeadersVisible = true;
          //  textBox2.Text = textBox2.Text + "kek";

            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
            columnHeaderStyle.BackColor = Color.AliceBlue;
            columnHeaderStyle.Font = new Font("Miriam Libre", 12, FontStyle.Regular);
            columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            Base10fromto26 kek = new Base10fromto26();
            for (int i = 0; i<columnsNum; i++)
            {
             
                string name =  kek.From10To26(i);
               // textBox2.Text = textBox2.Text + name;
                dataGridView1.Columns.Add(name, name);
              //  dataGridView1.Columns[i].SortMode = 
            }
            dataGridView1.RowHeadersVisible = true;
            dataGridView1.RowCount = rowsNum;

            DataGridViewCellStyle rowHeaderStyle = new DataGridViewCellStyle();
            rowHeaderStyle.BackColor = Color.Beige;
            rowHeaderStyle.Font = new Font("Miriam Libre", 12, FontStyle.Regular);

            dataGridView1.RowHeadersDefaultCellStyle = rowHeaderStyle;

            
                for (int i = 0; i<rowsNum; i++)
            {
                //textBox2.Text = textBox2.Text + (i+1).ToString();
                dataGridView1.Rows[i].HeaderCell.Value = (i).ToString();
            }
            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            dataGridView1.RowHeadersWidth = 70;
        }
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            /*
            try
            {
                int rowNum;
                string cell, columnNum;
                {
                    rowNum = e.RowIndex;
                    columnNum = f.From10To26(e.ColumnIndex);
                    cell = columnNum.ToString() + (rowNum).ToString();
                }
                if (dataGridView1.CurrentCell != null)
                {
                    myHash.AddExpression(cell, dataGridView1.CurrentCell.EditedFormattedValue.ToString());
                    myHash.ReCalculate(cell, converter);
                    ReWrite();
                }
            }
            catch (Exception)
            {

            }*/
        }

       /* private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                string name = "A.0";
                if (MyTable.GetInstance().expressions[name].ToString() != "")
                {
                    MyTable.GetInstance().expressions[name] = textBox1.Text.Substring(1);
                    myHash.ReCalculate(name, ref isRecalculated, converter);
                    ReWrite(ref isRecalculated);

                }

            }
        }*/
        public void ReWrite()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = "";
                }
            }
            foreach (DictionaryEntry pair in myHash.Values)
            {
                string index = pair.Key.ToString();
                int last = 0;
                while (index[last] >= 'A' && index[last] <= 'Z')
                {
                    last++;
                }
                string columnNum = index.Substring(0, last);
                string rowNum = index.Substring(last, index.Length - last);
                dataGridView1.Rows[Convert.ToInt32(rowNum)].Cells[f.From26To10(columnNum)].Value = pair.Value;
            }
        }
        public void Clear()
        {
           /* foreach (DictionaryEntry pair in myHash.Values)
            {
                string index = pair.Key.ToString();
                if (index.Split('.').Length == 2)
                {
                    string columnNum = index.Split('.')[0];
                    string rowNum = index.Split('.')[1];
                dataGridView1.Rows[Convert.ToInt32(rowNum)].Cells[f.From26To10(columnNum)].Value = "";
                }

            }
            myHash.values.Clear();
            myHash.expressions.Clear();
        */
          }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            
            string name = f.From10To26(e.ColumnIndex) + e.RowIndex;

            if (MyTable.GetInstance().expressions.Contains(name))
            {
                textBox1.Text = MyTable.GetInstance().expressions[name].ToString();
            }
            else
            {
                textBox1.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
             
        }
        public info openedForm = null;
        private void button7_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms.Count == 1)
            {
                openedForm = new info();
                openedForm.Show();
            }
            else
            {
                openedForm.Show();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

            if (dataGridView1.CurrentCell != null)
            {
                try
                {
                    int rowNum;
                    string cell, columnNum;
                    {
                        rowNum = dataGridView1.CurrentCell.RowIndex;
                        columnNum = f.From10To26(dataGridView1.CurrentCell.ColumnIndex);
                        cell = columnNum.ToString() + (rowNum).ToString();
                    }
                    myHash.AddExpression(cell, textBox1.Text);
                    myHash.ReCalculate(cell, converter);
                    ReWrite();
                }
                catch (Exception)
                {

                }
            }

        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            textBox1.ReadOnly = false;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            textBox1.ReadOnly = true;
        }
     
        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            string filetext = "";
            foreach (DictionaryEntry pair in myHash.Expressions)
            {

                string index = pair.Key.ToString();
                filetext = filetext + index + ' ' + pair.Value + "\n";
            }
            if (filetext.Length <= 4)
            {
                string message1 = "Saving failed!";
                string caption1 = "Alert";
                MessageBoxButtons buttons1 = MessageBoxButtons.OK;
                MessageBox.Show(message1, caption1, buttons1);
            }
            else
            {
                Form4 entername = new Form4(myHash);
                entername.Show();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form newForm = new Form2(this);
            newForm.Show();
           
        }

        private void button2_Click(object sender, EventArgs e) //добавление строки
        {
            string curRowNum = dataGridView1.CurrentCell.RowIndex.ToString();
            string curColumnNum = dataGridView1.CurrentCell.ColumnIndex.ToString();
            int curRowNumInt = Convert.ToInt32(curRowNum);
            int curColumnNumInt = Convert.ToInt32(curColumnNum);
            dataGridView1.Rows.Add();

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //textBox2.Text = textBox2.Text + (i+1).ToString();
                dataGridView1.Rows[i].HeaderCell.Value = (i).ToString();
            }
            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            dataGridView1.RowHeadersWidth = 70;
            var map = new Dictionary<string, string>();
            foreach (DictionaryEntry pair in myHash.Expressions)
            {
                string index = pair.Key.ToString();
                map[index] = pair.Value.ToString();

            }
            /*foreach (var pair in map)
            {
                string index = pair.Key;
                if (index.Split('.').Length == 2)
                {
                    string rowNum = index.Split('.')[0];
                    string columnNum = index.Split('.')[1];
                    int columnNumInt = Convert.ToInt32(columnNum);
                    if (columnNumInt)
                }
            }*/
            myHash.expressions.Clear();
            myHash.values.Clear();
            /* foreach (var pair in map)
             {
                 string index = pair.Key.ToString();
                 myHash.expressions[index] = "0";
                // myHash.ReCalculate(index, converter);
                 ReWrite();
             }*/
            foreach (var pair in map)
            {
                string index = pair.Key.ToString();
                string pairRowNum = "";
                string pairColumnNum = "";
                int j = 0;
                while (j < index.Length && index[j] >= 'A' && index[j] <= 'Z')
                {
                    pairColumnNum = pairColumnNum + index[j];
                    j++;
                } 
                while (j < index.Length && index[j] >= '0' && index[j] <= '9')
                {
                    pairRowNum = pairRowNum + index[j];
                    j++;
                }
                int pairRowNumInt = Convert.ToInt32(pairRowNum);
                string s = pair.Value.ToString();
                string new_s = "";
                for (int i = 0; i<s.Length; i++)
                {
                    if (s[i] >='A' && s[i] <='Z')
                    {

                        string rowNum = "";
                        string columnNum = ""; 
                        while (i<s.Length && s[i] >= 'A' && s[i] <= 'Z' )
                        {
                            columnNum = columnNum + s[i];
                            i++;
                        }
                        while (i<s.Length && s[i]>='0' && s[i] <= '9')
                        {
                            rowNum = rowNum + s[i];
                            i++;
                        }
                        i--;
                        int rowNumInt = Convert.ToInt32(rowNum);
                        if (rowNumInt >= curRowNumInt)
                        {
                            rowNum = (rowNumInt + 1).ToString();
                        }
                        new_s = new_s + columnNum + rowNum;

                    }
                    else
                    {
                        new_s = new_s + s[i];
                    }
                }
                if (pairRowNumInt >= curRowNumInt)
                {
                    index = pairColumnNum + (pairRowNumInt + 1).ToString();
                }
                if (new_s != "")
                {
                    myHash.AddExpression(index, new_s);
                }
            }
            //  myHash.ReCalculate("A0", converter);
            myHash.ReCalculateAll(converter);
            ReWrite();

        }

        private void button4_Click_1(object sender, EventArgs e) //удалить строчку
        {
            string curRowNum = dataGridView1.CurrentCell.RowIndex.ToString();
            string curColumnNum = dataGridView1.CurrentCell.ColumnIndex.ToString();
            int curRowNumInt = Convert.ToInt32(curRowNum);
            int curColumnNumInt = Convert.ToInt32(curColumnNum);
          //  dataGridView1.Rows.Add();

            var map = new Dictionary<string, string>();
            foreach (DictionaryEntry pair in myHash.Expressions)
            {
                string index = pair.Key.ToString();
                map[index] = pair.Value.ToString();

            }
            bool prv = false;
            foreach (var pair in map)
            {
                string index = pair.Key.ToString();
                string pairRowNum = "";
                string pairColumnNum = "";
                int j = 0;
                while (j < index.Length && index[j] >= 'A' && index[j] <= 'Z')
                {
                    pairColumnNum = pairColumnNum + index[j];
                    j++;
                }
                while (j < index.Length && index[j] >= '0' && index[j] <= '9')
                {
                    pairRowNum = pairRowNum + index[j];
                    j++;
                }
                int pairRowNumInt = Convert.ToInt32(pairRowNum);
                if (pairRowNumInt == curRowNumInt && pair.Value != "") prv = true;
            }
            if (prv)
            {
                string message1 = "Careful, there is data in the row. Delete anyway?";
                string caption1 = "Alert";
                MessageBoxButtons buttons1 = MessageBoxButtons.YesNo;
                if ( MessageBox.Show(message1, caption1, buttons1) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }
            myHash.expressions.Clear();
            myHash.values.Clear();

            foreach (var pair in map)
            {
                string index = pair.Key.ToString();
                string pairRowNum = "";
                string pairColumnNum = "";
                int j = 0;
                while (j < index.Length && index[j] >= 'A' && index[j] <= 'Z')
                {
                    pairColumnNum = pairColumnNum + index[j];
                    j++;
                } 
                while (j < index.Length && index[j] >= '0' && index[j] <= '9')
                {
                    pairRowNum = pairRowNum + index[j];
                    j++;
                }
                int pairRowNumInt = Convert.ToInt32(pairRowNum);
                string s = pair.Value.ToString();
                string new_s = "";
                for (int i = 0; i<s.Length; i++)
                {
                    if (s[i] >='A' && s[i] <='Z')
                    {

                        string rowNum = "";
                        string columnNum = ""; 
                        while (i<s.Length && s[i] >= 'A' && s[i] <= 'Z' )
                        {
                            columnNum = columnNum + s[i];
                            i++;
                        }
                        while (i<s.Length && s[i]>='0' && s[i] <= '9')
                        {
                            rowNum = rowNum + s[i];
                            i++;
                        }
                        i--;
                        int rowNumInt = Convert.ToInt32(rowNum);
                        if (rowNumInt == curRowNumInt)
                        {
                            rowNum = "0";
                        }
                        if (rowNumInt > curRowNumInt)
                        {
                            rowNum = (rowNumInt - 1).ToString();
                        }
                        new_s = new_s + columnNum + rowNum;

                    }
                    else
                    {
                        new_s = new_s + s[i];
                    }
                }
                if (pairRowNumInt == curRowNumInt) new_s = "";
                if (pairRowNumInt > curRowNumInt)
                {
                    index = pairColumnNum + (pairRowNumInt - 1).ToString();
                }
                if (new_s != "")
                {
                    myHash.AddExpression(index, new_s);
                }
            }
            dataGridView1.Rows.Remove(dataGridView1.Rows[curRowNumInt]);
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //textBox2.Text = textBox2.Text + (i+1).ToString();
                dataGridView1.Rows[i].HeaderCell.Value = (i).ToString();
            }
            //  myHash.ReCalculate("A0", converter);
            myHash.ReCalculateAll(converter);
            ReWrite();
        }

        private void button3_Click_1(object sender, EventArgs e) //добавление колонки
        {
            Base10fromto26 kek = new Base10fromto26();
            string curRowNum = dataGridView1.CurrentCell.RowIndex.ToString();
            string curColumnNum = dataGridView1.CurrentCell.ColumnIndex.ToString();
            int curRowNumInt = Convert.ToInt32(curRowNum);
            int curColumnNumInt = Convert.ToInt32(curColumnNum);
            string name = kek.From10To26(dataGridView1.Columns.Count);
            dataGridView1.Columns.Add(name,name);
           
            
            var map = new Dictionary<string, string>();
            foreach (DictionaryEntry pair in myHash.Expressions)
            {
                string index = pair.Key.ToString();
                map[index] = pair.Value.ToString();

            }
            myHash.expressions.Clear();
            myHash.values.Clear();
          
            foreach (var pair in map)
            {
                string index = pair.Key.ToString();
                string pairRowNum = "";
                string pairColumnNum = "";
                int j = 0;
                while (j < index.Length && index[j] >= 'A' && index[j] <= 'Z')
                {
                    pairColumnNum = pairColumnNum + index[j];
                    j++;
                }
                while (j < index.Length && index[j] >= '0' && index[j] <= '9')
                {
                    pairRowNum = pairRowNum + index[j];
                    j++;
                }
                int pairColumnNumInt = kek.From26To10(pairColumnNum);
                string s = pair.Value.ToString();
                string new_s = "";
                for (int i = 0; i < s.Length; i++)
                {
                    if (s[i] >= 'A' && s[i] <= 'Z')
                    {

                        string rowNum = "";
                        string columnNum = "";
                        while (i < s.Length && s[i] >= 'A' && s[i] <= 'Z')
                        {
                            columnNum = columnNum + s[i];
                            i++;
                        }
                        while (i < s.Length && s[i] >= '0' && s[i] <= '9')
                        {
                            rowNum = rowNum + s[i];
                            i++;
                        }
                        i--;
                        int columnNumInt = kek.From26To10(columnNum);
                        if (columnNumInt >= curColumnNumInt)
                        {
                            columnNum = kek.From10To26(columnNumInt + 1);
                        }
                        new_s = new_s + columnNum + rowNum;

                    }
                    else
                    {
                        new_s = new_s + s[i];
                    }
                }
                if (pairColumnNumInt >= curColumnNumInt)
                {
                    index =  kek.From10To26(pairColumnNumInt + 1) + pairRowNum;
                }
                if (new_s != "")
                {
                    myHash.AddExpression(index, new_s);
                }
            }
            myHash.ReCalculateAll(converter);
            ReWrite();
           
        }
        private void button1_Click_1(object sender, EventArgs e) // удалить колонку
        {
            string curRowNum = dataGridView1.CurrentCell.RowIndex.ToString();
            string curColumnNum = dataGridView1.CurrentCell.ColumnIndex.ToString();
            int curRowNumInt = Convert.ToInt32(curRowNum);
            int curColumnNumInt = Convert.ToInt32(curColumnNum);

            var map = new Dictionary<string, string>();
            foreach (DictionaryEntry pair in myHash.Expressions)
            {
                string index = pair.Key.ToString();
                map[index] = pair.Value.ToString();

            }
            Base10fromto26 kek = new Base10fromto26();
            bool prv = false;
            foreach (var pair in map)
            {
                string index = pair.Key.ToString();
                string pairRowNum = "";
                string pairColumnNum = "";
                int j = 0;
                while (j < index.Length && index[j] >= 'A' && index[j] <= 'Z')
                {
                    pairColumnNum = pairColumnNum + index[j];
                    j++;
                }
                while (j < index.Length && index[j] >= '0' && index[j] <= '9')
                {
                    pairRowNum = pairRowNum + index[j];
                    j++;
                }
                int pairColumnNumInt = kek.From26To10(pairColumnNum);
                if (pairColumnNumInt == curColumnNumInt && pair.Value != "") prv = true;
            }
            if (prv)
            {
                string message1 = "Careful, there is data in the row. Delete anyway?";
                string caption1 = "Alert";
                MessageBoxButtons buttons1 = MessageBoxButtons.YesNo;
                if (MessageBox.Show(message1, caption1, buttons1) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }
            myHash.expressions.Clear();
            myHash.values.Clear();

            foreach (var pair in map)
            {
                string index = pair.Key.ToString();
                string pairRowNum = "";
                string pairColumnNum = "";
                int j = 0;
                while (j < index.Length && index[j] >= 'A' && index[j] <= 'Z')
                {
                    pairColumnNum = pairColumnNum + index[j];
                    j++;
                }
                while (j < index.Length && index[j] >= '0' && index[j] <= '9')
                {
                    pairRowNum = pairRowNum + index[j];
                    j++;
                }
                int pairColumnNumInt = kek.From26To10(pairColumnNum);
                string s = pair.Value.ToString();
                string new_s = "";
                for (int i = 0; i < s.Length; i++)
                {
                    if (s[i] >= 'A' && s[i] <= 'Z')
                    {

                        string rowNum = "";
                        string columnNum = "";
                        while (i < s.Length && s[i] >= 'A' && s[i] <= 'Z')
                        {
                            columnNum = columnNum + s[i];
                            i++;
                        }
                        while (i < s.Length && s[i] >= '0' && s[i] <= '9')
                        {
                            rowNum = rowNum + s[i];
                            i++;
                        }
                        i--;
                        int columnNumInt = kek.From26To10(columnNum);
                        if (columnNumInt == curColumnNumInt)
                        {
                            columnNum = "Z";
                        }
                        if (columnNumInt > curColumnNumInt)
                        {
                            columnNum = kek.From10To26(columnNumInt - 1);
                        }
                        new_s = new_s + columnNum + rowNum;

                    }
                    else
                    {
                        new_s = new_s + s[i];
                    }
                }
                if (pairColumnNumInt == curColumnNumInt)
                {
                    new_s = "";
                }
                if (pairColumnNumInt > curColumnNumInt)
                {
                    index = kek.From10To26(pairColumnNumInt-1) + pairRowNum;
                }
                if (new_s != "")
                {
                    myHash.AddExpression(index, new_s);
                }
            }
            dataGridView1.Columns.Remove(dataGridView1.Columns[curColumnNumInt]);
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                //textBox2.Text = textBox2.Text + (i+1).ToString();
                dataGridView1.Columns[i].HeaderCell.Value = kek.From10To26(i);
            }
            //  myHash.ReCalculate("A0", converter);
            
            myHash.ReCalculateAll(converter);
            ReWrite();
        }
       

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string message = "Unsaved data will be lost. Are you sure you want to exit?";
            string caption = "Alert";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            if (MessageBox.Show(message, caption, buttons) == System.Windows.Forms.DialogResult.No)
            {
                e.Cancel = true;

            }
        }
        public void StringToTable (string s)
        {
            myHash.expressions.Clear();
            myHash.Values.Clear();
            string[] lines = s.Split('\n');
            for (int i = 0; i<lines.Length; i++)
            {
                string[] curLine = lines[i].Split(' ');
                string cell = curLine[0];
                string kek = lines[i];
                if (cell.Length+1 < lines[i].Length)
                {
                    string cellExpression = lines[i].Substring(cell.Length+1);
                    //  myHash.expressions[curLine[0]] = lines[i].Substring(curLine[0].Length);
                    myHash.AddExpression(cell, cellExpression);
                    myHash.ReCalculate(cell, converter);
                      ReWrite();
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Clear();
        }

    }
}
