using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

using System.Xml.Linq;
using Microsoft.Office.Core;

//"После выбора папки PKE программа должна построить список проверок, которые содержаться в папке PKE"

namespace VasilyevAV
{
    public partial class FormVasilyevAV : Form
    {
        //какие файлы парсим
        string[] fileExtChecks = {".xml", ".pke"};

        //схемы проверки
        enum t_pke_cxema {cxema1, cxema2, error};

        //блок RM3_ПКЭ
        struct t_RM3_PKE
        {
            public string ver;
            public string uid;
        }

        //данные проверки
        struct t_Check 
        {
            public t_RM3_PKE rmz;
            public ListViewItem itemCheck;
            public Dictionary<int, ListViewItem> itemsCheckDetails; //№пп, ListViewItem
            public t_pke_cxema pke_cxema()
            {
                const int cxemaIndex = 3;
                try
                {
                    if (itemCheck.SubItems[cxemaIndex].Text == "1")
                        return (t_pke_cxema.cxema1);
                    if (itemCheck.SubItems[cxemaIndex].Text == "2")
                        return (t_pke_cxema.cxema2);
                    return (t_pke_cxema.error);
                }
                catch
                {
                    return (t_pke_cxema.error);
                }
            }
        }

        //словарь проверок
        Dictionary<string, t_Check> checks = new Dictionary<string, t_Check>(); //UID, даные проверки
        
        //парсинг файла
        private bool ParseFile(string ffile)
        {
            try
            {
                if (!File.Exists(ffile))
                    return (false);
                if (0 == (new FileInfo(ffile)).Length)
                    return (false);

                t_Check check;
                UInt64 tempUInt64;
                ListViewItem lvi;
                XDocument doc = XDocument.Load(ffile);

                //если UID уникален, то новое измерение
                if (!checks.TryGetValue(doc.Root.Attribute("UID").Value, out check))
                {
                    ////////////////////////
                    //читаем RM3
                    check.rmz.ver = doc.Root.Attribute("Ver").Value;
                    check.rmz.uid = doc.Root.Attribute("UID").Value;

                    ////////////////////////
                    //читаем Param_Check_PKE
                    lvi = new ListViewItem();
                    //nameObject
                    lvi.Text = doc.Root.Element("Param_Check_PKE").Attribute("nameObject").Value;
                    //timeStart
                    if (!UInt64.TryParse(doc.Root.Element("Param_Check_PKE").Attribute("TimeStart").Value, out tempUInt64)) return (false);
                    lvi.SubItems.Add((new DateTime(1970, 1, 1, 0, 0, 0, 0)).AddMilliseconds(tempUInt64).ToString("dd/MM/yy hh:mm"));
                    //timeStop
                    if (!UInt64.TryParse(doc.Root.Element("Param_Check_PKE").Attribute("TimeStop").Value, out tempUInt64)) return (false);
                    lvi.SubItems.Add((new DateTime(1970, 1, 1, 0, 0, 0, 0)).AddMilliseconds(tempUInt64).ToString("dd/MM/yy hh:mm"));
                    //active_cxema
                    lvi.SubItems.Add(doc.Root.Element("Param_Check_PKE").Attribute("active_cxema").Value);
                    //averaging_interval_time
                    if (!UInt64.TryParse(doc.Root.Element("Param_Check_PKE").Attribute("averaging_interval_time").Value, out tempUInt64)) return (false);
                    if (tempUInt64 > (1000 * 60))
                        lvi.SubItems.Add(((double)tempUInt64 / (1000 * 60)).ToString(".00") + " мин");
                    else
                        if (tempUInt64 > (1000))
                            lvi.SubItems.Add(((double)tempUInt64 / (1000)).ToString(".00") + " c");
                        else
                            lvi.SubItems.Add(tempUInt64.ToString() + " мc");
                    check.itemCheck = lvi;
                    check.itemsCheckDetails = new Dictionary<int, ListViewItem>();
                    checks.Add(check.rmz.uid, check);
                }

                ////////////////////////
                //читаем Result_Check_PKE
                lvi = new ListViewItem();
                //timeTek
                if (!UInt64.TryParse(doc.Root.Element("Result_Check_PKE").Attribute("TimeTek").Value, out tempUInt64)) return (false);
                lvi.Text = (new DateTime(1970, 1, 1, 0, 0, 0, 0)).AddMilliseconds(tempUInt64).ToString("dd/MM/yy hh:mm");
                //при одинаковом UID: во всех файлах все pke_cxema в Param_Check_PKE и Result_Check_PKE должны быть одинаковы
                switch (doc.Root.Element("Result_Check_PKE").Attribute("pke_cxema").Value)
                {
                    case "1":
                        if (check.pke_cxema() != t_pke_cxema.cxema1) return (false);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("UA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("PA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("QA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("SA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("Freq").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("sigmaUy").Value);
                        break;
                    case "2":
                        if (check.pke_cxema() != t_pke_cxema.cxema2) return (false);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("UAB").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("UBC").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("UCA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IAB").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IBC").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("ICA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IA").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IB").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IC").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("PO").Value);

                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("PP").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("QO").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("QP").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("SO").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("SP").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("UO").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("UP").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IO").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("IP").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("KO").Value);

                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("Freq").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("sigmaUy").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("sigmaUyAB").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("sigmaUyBC").Value);
                        lvi.SubItems.Add(doc.Root.Element("Result_Check_PKE").Attribute("sigmaUyCA").Value);
                        break;
                    default:
                        return (false);
                }
                checks[check.rmz.uid].itemsCheckDetails.Add(checks[check.rmz.uid].itemsCheckDetails.Count(), lvi);
                return (true);
            }
            catch
            {
                MessageBox.Show("Ошибка при разборе файла!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (false);
            }
        }

        public FormVasilyevAV()
        {
            InitializeComponent();
        }

        private void выбратьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
                try
                {
                    checks.Clear();
                    listViewCheks.Items.Clear();
                    listViewCheckDetails_Cxema1.Visible = false;
                    listViewCheckDetails_Cxema2.Visible = false;
                    выйтиToolStripMenuItem.Enabled = false;
                    //парсим файлы
                    foreach (string ffile in Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.*", SearchOption.AllDirectories))
                        foreach (string cch in fileExtChecks)
                            if (cch == Path.GetExtension(ffile).ToLower().Trim())
                            {
                                ParseFile(ffile);
                                break;
                            }
                    //выводим в таблицу
                    foreach (KeyValuePair<string, t_Check> kvp in checks)
                        listViewCheks.Items.Add(kvp.Value.itemCheck);
                }
                catch
                {
                    MessageBox.Show("Ошибка при разборе каталога " + Path.GetFileName(folderBrowserDialog1.SelectedPath) + "!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
        }
        
        private void listViewCheks_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (0 == listViewCheks.SelectedItems.Count)
                {
                    listViewCheckDetails_Cxema1.Visible = false;
                    listViewCheckDetails_Cxema2.Visible = false;
                    выйтиToolStripMenuItem.Enabled = false;
                }
                else
                {
                    выйтиToolStripMenuItem.Enabled = true;
                    foreach (KeyValuePair<string, t_Check> kvp in checks)
                        if (kvp.Value.itemCheck == listViewCheks.SelectedItems[0])
                            switch (kvp.Value.pke_cxema())
                            {
                                case t_pke_cxema.cxema1:
                                    listViewCheckDetails_Cxema1.Items.Clear();
                                    listViewCheckDetails_Cxema1.Visible = true;
                                    listViewCheckDetails_Cxema2.Visible = false;
                                    foreach (KeyValuePair<int, ListViewItem> k in kvp.Value.itemsCheckDetails)
                                        listViewCheckDetails_Cxema1.Items.Add(k.Value);
                                    break;
                                case t_pke_cxema.cxema2:
                                    listViewCheckDetails_Cxema2.Items.Clear();
                                    listViewCheckDetails_Cxema2.Visible = true;
                                    listViewCheckDetails_Cxema1.Visible = false;
                                    foreach (KeyValuePair<int, ListViewItem> k in kvp.Value.itemsCheckDetails)
                                        listViewCheckDetails_Cxema2.Items.Add(k.Value);
                                    break;
                                default:
                                    return;
                            }
                }
            }
            catch
            {
                MessageBox.Show("Ошибка при разборе проверки!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void выйтиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ListView lv = listViewCheckDetails_Cxema1; ;
            foreach (KeyValuePair<string, t_Check> kvp in checks)
                if (kvp.Value.itemCheck == listViewCheks.SelectedItems[0])
                    switch (kvp.Value.pke_cxema())
                    {
                        case t_pke_cxema.cxema1:
                            lv = listViewCheckDetails_Cxema1;
                            break;
                        case t_pke_cxema.cxema2:
                            lv = listViewCheckDetails_Cxema2;
                            break;
                        default:
                            return;
                    }
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workBook;
            Excel.Worksheet workSheet;
            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            int currentCol = 1;
            int currentRow = 1;
            workSheet.Rows[currentRow].NumberFormat = "@";
            foreach (ColumnHeader ch in lv.Columns)
                workSheet.Cells[currentRow, currentCol++] = ch.Text;
            foreach (ListViewItem lvi in lv.Items)
            {
                currentCol = 1;
                currentRow++;
                workSheet.Rows[currentRow].NumberFormat = "@";
                for (int ii = 0; ii < lvi.SubItems.Count; ii++)
                    workSheet.Cells[currentRow, currentCol++] = lvi.SubItems[ii].Text;
            }

        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
