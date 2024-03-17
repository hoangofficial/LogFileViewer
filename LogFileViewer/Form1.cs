using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using app = Microsoft.Office.Interop.Excel.Application;

namespace LogFileViewer
{
    public partial class frmMain : Form
    {
        public int target = -1;
        public frmMain()
        {
            InitializeComponent();
            dataGridView1.RowPrePaint += dataGridView1_RowPrePaint;
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
           
        }               
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            this.dataGridView1.Columns[0].Width = 35;
            this.dataGridView1.Columns[1].Width = 140;
            this.dataGridView1.Columns[2].Width = 70;
            this.dataGridView1.Columns[3].Width = 90;
            this.dataGridView1.Columns[4].Width = 560;
            
            string filePath = richTextBox1.Text;              
            if (File.Exists(filePath))
            {
                try
                {
                    string[] lines = File.ReadAllLines(filePath);
                    target = 0;
                    foreach (string line in lines)
                    {
                        if (line.Contains(richTextBox2.Text) && line.Contains("=-1") && line.Length > 160)                            
                        {
                            target = -1;
                            string[] parts = line.Split(' ');
                            string date = parts[0] + " " + parts[1];
                            string deb = parts[5];
                            string eqp = parts[8].Substring(4);                            
                            //
                            string startPattern = "RTN_CD=-1";
                            string endPattern = " ERR_MSG_LOC";

                            int startIndex = line.IndexOf(startPattern);
                            int endIndex = line.LastIndexOf(endPattern);

                            if (startIndex >= 0 && endIndex >= 0 && endIndex > startIndex)
                            {
                                string result = line.Substring(startIndex, endIndex - startIndex);                                
                                dataGridView1.Rows.Add("", date, deb, eqp, result);
                            }                            
                        }
                    }
                    if (target !=-1)
                    {
                        MessageBox.Show("Ther is no error in this time.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading log file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("File not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            richTextBox2.Text = "";
            richTextBox1.Text = "";
        }
        //Tao event RowPrePaint de danh STT
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (!e.RowIndex.Equals(dataGridView1.NewRowIndex))
            {
                dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
            }
        }        
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.Multiselect = true; Do not allow to multiselect
            openFileDialog.Filter = "Log Files (*.log)|*.log";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] selectedFiles = openFileDialog.FileNames;
                //foreach (string file in selectedFiles)
                //{                    
                //    richTextBox1.AppendText(file + Environment.NewLine);
                //}
                
                
                richTextBox1.Text = openFileDialog.FileName;
            }            
        }

        private void xuatfile_excel(DataGridView g, string filename)
        {
            app obj = new app();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;

            for (int i = 1; i < g.Columns.Count + 1; i++)
            {
                obj.Cells[1, i] = g.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < g.Rows.Count; i++)
            {
                for (int j = 0; j < g.Columns.Count; j++)
                {
                    if (g.Rows[i].Cells[j].Value != null)
                    {
                        obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            obj.ActiveWorkbook.SaveAs(filename);
            obj.ActiveWorkbook.Saved = true;
            obj.Quit();
            MessageBox.Show("Export successfully!");
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel Workbook|*.xlsx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                xuatfile_excel(dataGridView1, saveFileDialog1.FileName);
            }
        }
    }
}
