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
using System.Globalization;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WQT_Welders
{
    public partial class MainForm : Form
    {


        public static DataTable dtWelder = new DataTable();
        public static DataTable dtJoints = new DataTable();
        frmLoading newload = new frmLoading();

        public MainForm()
        {
            InitializeComponent();
            backgroundWorker1.RunWorkerAsync();
           
            InitializeData.InitWelders(dtWelder);
            InitializeData.InitJoints(dtJoints);
        }

        private void btnLoadWelder_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
            clearComponents();

            DataRow[] welders;

            //FOR ALL WELDERS COUNT LABEL
            if (cmbSubCon.SelectedItem.ToString().Equals("ALL"))
            {
                welders = dtWelder.Select();
            }
            else if (cmbSubCon.SelectedItem.ToString().Equals("NSH"))
            { 
                welders = dtWelder.Select("Subcontractor in ('NSH','NSH2')");
            }
            else  
            {
                welders = dtWelder.Select("Subcontractor='" + cmbSubCon.SelectedItem.ToString() + "'");
            }

            CheckWelderStatus(welders);
            lbl_welders.Text = welders.Count().ToString();


            //FOR ACTIVE WELDERS COUNT LABEL
            DataRow[] cnt_active = new DataRow[]{};
            if (cmbSubCon.SelectedItem.ToString().Equals("ALL"))
            {
                cnt_active = dtWelder.Select("ACTIVE='True'");
            }
            else if (cmbSubCon.SelectedItem.ToString().Equals("NSH"))
            {
                cnt_active = dtWelder.Select("Subcontractor in ('NSH','NSH2') AND ACTIVE='True'");
            }
            else
            {
                cnt_active = dtWelder.Select("Subcontractor='" + cmbSubCon.SelectedItem.ToString() + "' AND Active='True'");
            }
            lbl_active.Text = cnt_active.Count().ToString();

            //FOR CANCELED WELDERS COUNT LABEL
            DataRow[] cnt_cancelled = new DataRow[] { };
            if (cmbSubCon.SelectedItem.ToString().Equals("ALL"))
            {
                cnt_cancelled = dtWelder.Select("ACTIVE='False'");
            }
            else if (cmbSubCon.SelectedItem.ToString().Equals("NSH"))
            {
                cnt_cancelled = dtWelder.Select("Subcontractor in ('NSH','NSH2') AND ACTIVE='False'");
            }
            else
            {
                cnt_cancelled = dtWelder.Select("Subcontractor='" + cmbSubCon.SelectedItem.ToString() + "' AND Active='False'");
            }
            lbl_notactive.Text = cnt_cancelled.Count().ToString();
        }

        public void CheckWelderStatus(DataRow[] w)
        {
            DataTable welderResult = new DataTable();
            string[] row = new string[] {};
            DataRow[] wqt, nowqt;
           
            int ctr_x = 0;

            for (int n = 0; n < w.Count(); n++)
            {
                wqt = dtJoints.Select("welder1 ='" + w[n].ItemArray[1].ToString().Trim() + "' AND welder2 ='" + w[n].ItemArray[1].ToString().Trim() + "' AND wqt = 1");
                nowqt = dtJoints.Select("welder1 ='" + w[n].ItemArray[1].ToString().Trim() + "' and welder2 ='" + w[n].ItemArray[1].ToString().Trim() + "' AND wqt = 0");

                var maxWqt = wqt.AsEnumerable()
                               .Select(cols => cols.Field<DateTime>("dateofweld"))
                               .OrderByDescending(p => p.Ticks)
                               .FirstOrDefault();

                var minNoWqt = nowqt.AsEnumerable()
                               .Select(cols => cols.Field<DateTime>("dateofweld"))
                               .OrderBy(p => p.Ticks)
                               .FirstOrDefault();

                var minWqt = wqt.AsEnumerable()
                               .Select(cols => cols.Field<DateTime>("dateofweld"))
                               .OrderBy(p => p.Ticks)
                               .FirstOrDefault();

                if (wqt.Count() < 10)
                {
                    row = new string[] { w[n].ItemArray[1].ToString().Trim()
                                        , "ON-GOING WQT"
                                        , minNoWqt < minWqt && minNoWqt.ToString() != "01/01/0001 12:00:00 AM" ? "1" : "0"
                                        , w[n].ItemArray[7].ToString() == "True" ? "ACTIVE" : "NOT ACTIVE"
                                        , w[n].ItemArray[3].ToString()
                                        , w[n].ItemArray[4].ToString()
                                        , w[n].ItemArray[5].ToString()
                                        , w[n].ItemArray[6].ToString() };
                }
                else
                {
                    row = new string[] { w[n].ItemArray[1].ToString().Trim()
                                            ,minNoWqt < maxWqt && minNoWqt.ToString() != "01/01/0001 12:00:00 AM" ? "X": "OK"
                                            ,minNoWqt < minWqt && minNoWqt.ToString() != "01/01/0001 12:00:00 AM" ? "1" : "0"
                                            , w[n].ItemArray[7].ToString() == "True" ? "ACTIVE" : "NOT ACTIVE"
                                            , w[n].ItemArray[3].ToString()
                                            , w[n].ItemArray[4].ToString()
                                            , w[n].ItemArray[5].ToString()
                                            , w[n].ItemArray[6].ToString() };

                    if (minNoWqt < maxWqt && minNoWqt.ToString() != "01/01/0001 12:00:00 AM")
                    { 
                        ctr_x += 1;
                    }
                }

                var welderdata = new ListViewItem(row);
                lv_Welders.Items.Add(welderdata);
            }

            lbl_x.Text = ctr_x.ToString();
        }

        private void lv_Welders_SelectedIndexChanged(object sender, EventArgs e)
        {
            lv_matrixA.Items.Clear();
            lv_matrixB.Items.Clear();
            rd_all.Checked = true;

            try
            {
                DataRow[] wqt, nowqt;
                string[] rowB = new string[] { };
                string[] rowA = new string[] { };

                string welder = "";
                
                if (lv_Welders.SelectedItems.Count == 0)
                {
                    return;
                }
                else
                {
                    welder = lv_Welders.SelectedItems[0].Text;
                }

                if(welder != "")
                {
                    wqt = dtJoints.Select("welder1='" + lv_Welders.SelectedItems[0].Text + "' AND welder2='" + lv_Welders.SelectedItems[0].Text + "' AND wqt = 1","DATEOFWELD ASC");
                    nowqt = dtJoints.Select("welder1='" + lv_Welders.SelectedItems[0].Text + "' and welder2='" + lv_Welders.SelectedItems[0].Text + "' AND wqt = 0","DATEOFWELD ASC");

                    int ctrWqt = 0, ctrNoWqt = 0;

                    for (int i = 0; i < wqt.Count(); i++ )
                    {
                        ctrWqt +=1;
                        rowA = new string[] { ctrWqt.ToString()
                                            ,wqt[i][0].ToString().Trim()
                                            ,wqt[i][1].ToString().Trim()
                                            ,Convert.ToDateTime(wqt[i][2].ToString()).ToString("yyyy/MM/dd")
                                            ,wqt[i][3].ToString().Trim()
                                            ,wqt[i][4].ToString().Trim()
                                            ,wqt[i][5].ToString().Trim()
                                            ,wqt[i][6].ToString().Trim()
                                            ,wqt[i][7].ToString().Trim()
                                            ,wqt[i][8].ToString().Trim().Equals("True") ? "YES" : "NO"
                                            ,wqt[i][9].ToString().Trim() 
                                            ,wqt[i][10].ToString().Trim() 
                                            ,wqt[i][11].ToString().Trim()
                                            ,wqt[i][12].ToString().Trim()
                                            ,wqt[i][13].ToString().Trim() 
                                            ,wqt[i][14].ToString().Trim()
                                            ,wqt[i][15].ToString().Trim()
                                            ,wqt[i][16].ToString().Trim()
                                            ,wqt[i][17].ToString().Trim()
                                    };
                        var dataA = new ListViewItem(rowA);
                        lv_matrixA.Items.Add(dataA);
                    }

                    for (int i = 0; i < nowqt.Count(); i++)
                    {
                        ctrNoWqt += 1;
                        rowB = new string[] { ctrNoWqt.ToString()
                                            , nowqt[i][0].ToString().Trim()
                                            , nowqt[i][1].ToString().Trim()
                                            , Convert.ToDateTime(nowqt[i][2].ToString()).ToString("yyyy/MM/dd")
                                            , nowqt[i][3].ToString().Trim()
                                            , nowqt[i][4].ToString().Trim()
                                            , nowqt[i][5].ToString().Trim()
                                            , nowqt[i][6].ToString().Trim()
                                            , nowqt[i][7].ToString().Trim()
                                            , nowqt[i][8].ToString().Trim().Equals("True") ? "YES" : "NO"
                                            , nowqt[i][9].ToString().Trim()
                                            , nowqt[i][10].ToString().Trim()
                                            , nowqt[i][11].ToString().Trim()
                                            , nowqt[i][12].ToString().Trim()
                                            , nowqt[i][13].ToString().Trim()
                                            , nowqt[i][14].ToString().Trim()
                                            , nowqt[i][15].ToString().Trim()
                                            , nowqt[i][16].ToString().Trim()
                                            , nowqt[i][17].ToString().Trim()
                                    };
                        var dataB = new ListViewItem(rowB);
                        lv_matrixB.Items.Add(dataB);
                    }
                }                
            }
            catch (Exception excp)
            {
                MessageBox.Show(excp.ToString());
            } 
        }


        public void LoadingForm()
        {
            if (newload != null)
            {
                newload = new frmLoading();
                newload.StartPosition = FormStartPosition.CenterScreen;
                newload.Location = new Point(this.Location.X + (Width - newload.Width) / 2, this.Location.Y + (Height - newload.Height) / 2);
                Application.Run(newload);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (!newload.IsHandleCreated)
            {
                Thread t = new Thread(new ThreadStart(LoadingForm));
                t.IsBackground = true;
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
            }
            else
            {
                this.Invoke((MethodInvoker)delegate()
                {
                    newload.TopLevel = true;
                    newload.Show();
                });
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (newload != null)
            {
                this.Invoke(new Action(() => newload.Close()));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
            ExporttoExcel();
        }

        public void ExporttoExcel()
        {
            Missing mv = Missing.Value;
            Microsoft.Office.Interop.Excel.Application NewExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook wkb = NewExcelApp.Workbooks.Add(mv);
            Excel.Worksheet wks1 = wkb.Sheets[1];
            
            Excel.Sheets addWks = null;
            Excel.Worksheet wks2 = null;
            Excel.Worksheet wks3 = null;

            addWks = wkb.Sheets as Excel.Sheets;
            wks2 = (Excel.Worksheet)addWks.Add(wks1, Type.Missing, Type.Missing, Type.Missing);
            wks3 = (Excel.Worksheet)addWks.Add(wks1, Type.Missing, Type.Missing, Type.Missing);

            wks1.Name = "WELDERS -"+ cmbSubCon.SelectedItem.ToString();
            wks2.Name = "WQT JOINTS" ;
            wks3.Name = "NON-WQT JOINTS"; 

            
            //WELDERS
            for (int y = 0; y <= lv_Welders.Columns.Count - 1; y++)
            {
                ((Excel.Range)wks1.Cells[1, y + 1]).Value = lv_Welders.Columns[y].Text;
            }

            for (int x = 1; x <= lv_Welders.Items.Count; x++)
            {
                for (int y = 0; y <= lv_Welders.Columns.Count - 1; y++)
                {
                    ((Excel.Range)wks1.Cells[x + 1, y + 1]).NumberFormat = "@";
                    ((Excel.Range)wks1.Cells[x + 1, y + 1]).Value = lv_Welders.Items[x - 1].SubItems[y].Text;
                }
            }

            //WQT JOINTS
            for (int y = 0; y <= lv_matrixA.Columns.Count - 1; y++)
            {
                ((Excel.Range)wks2.Cells[1, y + 1]).Value = lv_matrixA.Columns[y].Text;
            }

            for (int x = 1; x < lv_matrixA.Items.Count; x++)
            {
                for (int y = 0; y < lv_matrixA.Columns.Count - 1; y++)
                {
                    ((Excel.Range)wks2.Cells[x + 1, y + 1]).NumberFormat = "@";
                    ((Excel.Range)wks2.Cells[x + 1, y + 1]).Value = lv_matrixA.Items[x - 1].SubItems[y].Text;
                }
            }

            //NON-WQT JOINTS
            for (int y = 0; y <= lv_matrixB.Columns.Count - 1; y++)
            {
                ((Excel.Range)wks3.Cells[1, y + 1]).Value = lv_matrixB.Columns[y].Text;
            }

            for (int x = 1; x <= lv_matrixB.Items.Count; x++)
            {
                for (int y = 0; y <= lv_matrixB.Columns.Count - 1; y++)
                {
                    ((Excel.Range)wks3.Cells[x + 1, y + 1]).NumberFormat = "@";
                    ((Excel.Range)wks3.Cells[x + 1, y + 1]).Value = lv_matrixB.Items[x - 1].SubItems[y].Text;
                }
            }


            SaveFileDialog sfd = new SaveFileDialog();
            string filePath = "";

            if (sfd.ShowDialog(this) == DialogResult.OK)
            {
                filePath = sfd.FileName;
            }

            if (filePath != "")
            {
                wkb.SaveAs(filePath);
                NewExcelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(addWks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wks1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wks2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wks3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(NewExcelApp);
                MessageBox.Show("SUCCESSFULLY EXPORTED!", "EXPORT");
            }
         }

        private void lv_matrixA_ColumnReordered(object sender, ColumnReorderedEventArgs e)
        {
        }

        private void radiobtnLoadWelder_CheckedChanged(object sender, EventArgs e)
        {
            if (radiobtnLoadWelder.Checked)
            {
                lv_matrixB.Items.Clear();
                string[] rowB = new string[] { };
                int ctrNoWqt = 0;
                DataRow[] wqt, nowqt;


                wqt = dtJoints.Select("welder1='" + lv_Welders.SelectedItems[0].Text + "' AND welder2='" + lv_Welders.SelectedItems[0].Text + "' AND wqt = 1", "DATEOFWELD ASC");
                nowqt = dtJoints.Select("welder1='" + lv_Welders.SelectedItems[0].Text + "' and welder2='" + lv_Welders.SelectedItems[0].Text + "' AND wqt = 0", "DATEOFWELD ASC");

                var maxWqt = wqt.AsEnumerable()
                   .Select(cols => cols.Field<DateTime>("dateofweld")).Take(10)
                   .OrderByDescending(p => p.Ticks)
                   .FirstOrDefault();

                for (int i = 0; i < nowqt.Count(); i++)
                {
                    if (Convert.ToDateTime(nowqt[i][2].ToString()) < maxWqt)
                    {
                        ctrNoWqt += 1;
                        rowB = new string[] { 

                                            ctrNoWqt.ToString()
                                            , nowqt[i][0].ToString().Trim()
                                            , nowqt[i][1].ToString().Trim()
                                            , Convert.ToDateTime(nowqt[i][2].ToString()).ToString("yyyy/MM/dd")
                                            , nowqt[i][3].ToString().Trim()
                                            , nowqt[i][4].ToString().Trim()
                                            , nowqt[i][5].ToString().Trim()
                                            , nowqt[i][6].ToString().Trim()
                                            , nowqt[i][7].ToString().Trim()
                                            , nowqt[i][8].ToString().Trim().Equals("True") ? "YES" : "NO"
                                            , nowqt[i][9].ToString().Trim()
                                            , nowqt[i][10].ToString().Trim() 
                                            , nowqt[i][11].ToString().Trim()
                                            , nowqt[i][12].ToString().Trim() 
                                            , nowqt[i][13].ToString().Trim()
                                            , nowqt[i][14].ToString().Trim()
                                            , nowqt[i][15].ToString().Trim()
                                            , nowqt[i][16].ToString().Trim()
                                            , nowqt[i][17].ToString().Trim()
                                    };
                        var dataB = new ListViewItem(rowB);
                        lv_matrixB.Items.Add(dataB);
                    }
                }
            
            }
           
        }

        private void rd_all_CheckedChanged(object sender, EventArgs e)
        {
            if (rd_all.Checked)
            {
                lv_matrixB.Items.Clear();
                DataRow[] wqt, nowqt;
                string[] rowB = new string[] { };
                string[] rowA = new string[] { };

                string welder = "";

                if (lv_Welders.SelectedItems.Count == 0)
                {
                    return;
                }
                else
                {
                    welder = lv_Welders.SelectedItems[0].Text;
                }

                if (welder != "")
                {
                    wqt = dtJoints.Select("welder1='" + lv_Welders.SelectedItems[0].Text + "' AND welder2='" + lv_Welders.SelectedItems[0].Text + "' AND wqt = 1", "DATEOFWELD ASC");
                    nowqt = dtJoints.Select("welder1='" + lv_Welders.SelectedItems[0].Text + "' and welder2='" + lv_Welders.SelectedItems[0].Text + "' AND wqt = 0", "DATEOFWELD ASC");

                    int ctrNoWqt = 0;

                    for (int i = 0; i < nowqt.Count(); i++)
                    {
                        ctrNoWqt += 1;
                        rowB = new string[] { 
                                       
                                        ctrNoWqt.ToString()
                                            , nowqt[i][0].ToString().Trim()
                                            , nowqt[i][1].ToString().Trim()
                                            , Convert.ToDateTime(nowqt[i][2].ToString()).ToString("yyyy/MM/dd")
                                            , nowqt[i][3].ToString().Trim()
                                            , nowqt[i][4].ToString().Trim()
                                            , nowqt[i][5].ToString().Trim()
                                            , nowqt[i][6].ToString().Trim()
                                            , nowqt[i][7].ToString().Trim()
                                            , nowqt[i][8].ToString().Trim().Equals("True") ? "YES" : "NO"
                                            , nowqt[i][9].ToString().Trim()
                                            , nowqt[i][10].ToString().Trim() 
                                            , nowqt[i][11].ToString().Trim()
                                            , nowqt[i][12].ToString().Trim() 
                                            , nowqt[i][13].ToString().Trim()
                                            , nowqt[i][14].ToString().Trim()
                                            , nowqt[i][15].ToString().Trim()
                                            , nowqt[i][16].ToString().Trim()
                                            , nowqt[i][17].ToString().Trim()
                                    };
                        var dataB = new ListViewItem(rowB);
                        lv_matrixB.Items.Add(dataB);
                    }
                }
            }
              
        }

        private void lv_matrixA_KeyUp(object sender, KeyEventArgs e)
        {
            if (sender != lv_matrixA) return;

            if (e.Control && e.KeyCode == Keys.C)
                CopySelectedValuesToClipboard(lv_matrixA);
        }

        private void CopySelectedValuesToClipboard(ListView lv)
        {
            var builder = new StringBuilder();

            StringBuilder sb = new StringBuilder();
            foreach (var item in lv.SelectedItems)
            {
                ListViewItem l = item as ListViewItem;
                if (l != null)
                    foreach (ListViewItem.ListViewSubItem sub in l.SubItems)
                        sb.Append(sub.Text + "\t");
                sb.AppendLine();
            }
            Clipboard.SetDataObject(sb.ToString());
        }

        private void lv_matrixB_KeyUp(object sender, KeyEventArgs e)
        {
            if (sender != lv_matrixB) return;

            if (e.Control && e.KeyCode == Keys.C)
                CopySelectedValuesToClipboard(lv_matrixB);
        }

        private void lv_Welders_KeyUp(object sender, KeyEventArgs e)
        {
            if (sender != lv_Welders) return;

            if (e.Control && e.KeyCode == Keys.C)
                CopySelectedValuesToClipboard(lv_Welders);
        }

        private void lv_matrixA_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            dtJoints.Rows.Clear();
            InitializeData.InitJoints(dtJoints);
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            dtWelder.Rows.Clear();
            InitializeData.InitWelders(dtWelder);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
            clearComponents();
            dtWelder.Rows.Clear();
            InitializeData.InitWelders(dtWelder);

            dtJoints.Rows.Clear();
            InitializeData.InitJoints(dtJoints);
        }

        public void clearComponents()
        {
            lv_Welders.Items.Clear();
            lv_matrixA.Items.Clear();
            lv_matrixB.Items.Clear();
            lbl_active.Text = "0";
            lbl_notactive.Text = "0";
            lbl_welders.Text = "0";
            lbl_x.Text = "0";
        }

        private void backgroundWorker4_DoWork(object sender, DoWorkEventArgs e)
        {
            backgroundWorker2.RunWorkerAsync();
            backgroundWorker3.RunWorkerAsync();
        }
    }
}
