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
using System.Diagnostics;

namespace LogFileViewer
{
    public partial class Form2 : Form
    {
        public int target = -1;
        public Form2()
        {
            InitializeComponent();            
            dataGridView1.RowPrePaint += dataGridView1_RowPrePaint;
        }

        private void searchButton_Click(object sender, EventArgs e)
        {            
            dataGridView1.Rows.Clear();           
            this.dataGridView1.Columns[0].Width = 35;
            this.dataGridView1.Columns[1].Width = 140;
            this.dataGridView1.Columns[2].Width = 70;
            this.dataGridView1.Columns[3].Width = 90;
            this.dataGridView1.Columns[4].Width = 571;
            string searchQuery = txtShow.Text;
            string logFolderPath = txtShow.Text;
            DateTime startTime = dateTimePicker1.Value;
            DateTime endTime = dateTimePicker2.Value;
            List<DateTime> timeList = new List<DateTime>();
            while (startTime < endTime)
            {
                timeList.Add(startTime);
                startTime = startTime.AddHours(1);
            }
            foreach (DateTime time in timeList)
            {
                string[] logFiles = new string[] { time.ToString("yyyy-MM-dd_HH") + ".log" };                

                foreach (string logFile in logFiles)
                {
                    string logFilePath = Path.Combine(logFolderPath, logFile);

                    if (File.Exists(logFilePath))
                    {
                        txtLogStart.Text = dateTimePicker1.Value.ToString("yyyy-MM-dd_HH") + ".log";
                        txtLogEnd.Text = dateTimePicker2.Value.ToString("yyyy-MM-dd_HH") + ".log";
                        try
                        {
                            string[] lines = File.ReadAllLines(txtShow.Text + logFile);
                            target = 0;
                            foreach (string line in lines)
                            {
                                if (line.Contains(txtCondition.Text) && line.Contains("=-1") && line.Length > 160)
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
                            //if (target != -1)
                            //{
                            //    MessageBox.Show("Ther is no error in this time.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            //    break;
                            //}
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error reading log file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            foreach (DateTime time in timeList)
            {
                string[] logFiles = new string[] { "ECS.Equipment.Host.COPCDriver_Tib_M_"+ time.ToString("yyyy-MM-dd_HH") + ".log" };

                foreach (string logFile in logFiles)
                {
                    string logFilePath = Path.Combine(logFolderPath, logFile);

                    if (File.Exists(logFilePath))
                    {
                        txtLogStart.Text = "ECS.Equipment.Host.COPCDriver_Tib_M_" + dateTimePicker1.Value.ToString("yyyy-MM-dd_HH") + ".log";
                        txtLogEnd.Text = "ECS.Equipment.Host.COPCDriver_Tib_M_" + dateTimePicker2.Value.ToString("yyyy-MM-dd_HH") + ".log";
                        try
                        {
                            string[] lines = File.ReadAllLines(txtShow.Text + logFile);
                            target = 0;
                            foreach (string line in lines)
                            {
                                if (line.Contains(txtCondition.Text) && line.Contains("=-1") && line.Length > 160)
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
                            //if (target != -1)
                            //{
                            //    MessageBox.Show("Ther is no error in this time.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            //    break;
                            //}
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error reading log file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            //txtPath.Text = "";
            txtShow.Text = "";
            txtCondition.Text = "";
            txtLogStart.Text = "";
            txtLogEnd.Text = "";
            dataGridView1.Rows.Clear();
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Log Files (*.log)|*.log";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //txtShow.Text = openFileDialog.FileName;
                int lastIndex = openFileDialog.FileName.LastIndexOf('\\');

                if (lastIndex >= 0)
                {
                    // Lấy đường dẫn từ đầu đến dấu '\' cuối cùng
                    string folderPath = openFileDialog.FileName.Substring(0, lastIndex + 1);

                    // Hiển thị đường dẫn lên ô văn bản (txtPath)
                    txtShow.Text = folderPath;
                }
                else
                {
                    
                }
            }

            //FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            //DialogResult result = folderBrowser.ShowDialog();
            //if (result == DialogResult.OK)
            //{
            //    txtShow.Text = folderBrowser.SelectedPath + @"\";
            //}
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (!e.RowIndex.Equals(dataGridView1.NewRowIndex))
            {
                dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
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

        private void button1_Click(object sender, EventArgs e)
        {
            Process currentProcess = Process.GetCurrentProcess();
            ProcessThreadCollection threadCollection = currentProcess.Threads;
            MessageBox.Show($"Thread: {threadCollection.Count}");
        }
    }
}
