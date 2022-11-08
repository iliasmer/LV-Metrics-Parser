using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LV_Metrics_Parser
{
    public partial class ParseByUsageForm : Form
    {
        public ParseByUsageForm()
        {
            InitializeComponent();
        }

        List<ImportedData> dataSumDEIList = new List<ImportedData>();
        List<ImportedData> dataSumPKYList = new List<ImportedData>();
        string[] diktiaList = new string[] { "1", "2", "3", "4" };
        string[] XriseisList = new string[] {   "ΑΓΡΟΤΙΚΗ", 
                                                "ΑΓΡΟΤΙΚΗ ΠΚΥ", 
                                                "ΒΙΟΜΗΧΑΝΙΚΗ", 
                                                "ΒΙΟΜΗΧΑΝΙΚΗ ΠΚΥ",                         
                                                "ΓΕΝΙΚΗ", 
                                                "ΓΕΝΙΚΗ ΠΚΥ", 
                                                "ΔΗΜΟΣΙΑ", 
                                                "ΔΗΜΟΣΙΑ ΠΚΥ",
                                                "ΚΟΙΝΟΤΙΚΑ - ΦΟΠ", 
                                                "ΚΟΙΝΟΤΙΚΑ - ΦΟΠ ΠΚΥ", 
                                                "ΝΠΔΔ - Δ. ΕΠΙΧ. - ΠΡ", 
                                                "ΝΠΔΔ - Δ. ΕΠΙΧ. - ΠΡ ΠΚΥ",
                                                "ΟΙΚΙΑΚΗ", 
                                                "ΟΙΚΙΑΚΗ ΠΚΥ"};

        private void monthSelectbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            monthLabel.Text = monthSelectbox.SelectedIndex.ToString() + " " + monthSelectbox.Text;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            sourceFileLabelName.Text = "";
            importFileProgressLabel.Text = "";
            exportFileNameLabelDEI.Text = "";
            exportFileProgressLabelDEI.Text = "";
            exportFileNameLabelPKY.Text = "";
            exportFileProgressLabelPKY.Text = "";
        }

        private ImportedData nullifyDataSum (ImportedData sumOfCategory) {
            sumOfCategory.Minas = "0";
            sumOfCategory.Kata = "0";
            sumOfCategory.Sys = "0";
            sumOfCategory.Domi = "0";
            sumOfCategory.Xrisi = "0";
            sumOfCategory.Timologio = "0";
            sumOfCategory.ArLogar = "0";
            sumOfCategory.KWH_H = "0";
            sumOfCategory.KWH_N = "0";
            sumOfCategory.KWH_T = "0";
            sumOfCategory.Pagio = "0";
            sumOfCategory.ElXreosi = "0";
            sumOfCategory.Energia = "0";
            sumOfCategory.Isxys = "0";
            sumOfCategory.DEnergias = "0";
            sumOfCategory.CO2 = "0";
            sumOfCategory.EkptEt = "0";
            sumOfCategory.EkptOgk = "0";
            sumOfCategory.AntEkpt = "0";
            sumOfCategory.EkptTY = "0";
            sumOfCategory.TK_KY = "0";
            sumOfCategory.Mig = "0";
            sumOfCategory.EkptPag = "0";
            sumOfCategory.NeaEkptPagiou = "0";
            sumOfCategory.EkptPoso = "0";
            sumOfCategory.Ekpt8 = "0";
            sumOfCategory.EkptosiGP = "0";
            sumOfCategory.Ekpt10 = "0";
            sumOfCategory.Ekpt15 = "0";
            sumOfCategory.Pro = "0";
            sumOfCategory.EkpOgkDim = "0";
            sumOfCategory.Ekpt2 = "0";
            sumOfCategory.EkptEnXT = "0";
            sumOfCategory.Kouponi = "0";
            sumOfCategory.EkptEbill = "0";
            sumOfCategory.GreenPass = "0";
            sumOfCategory.EkGreenPass = "0";
            sumOfCategory.XXS = "0";
            sumOfCategory.XXD = "0";
            sumOfCategory.YKO = "0";
            sumOfCategory.LX = "0";
            sumOfCategory.ETMEAR = "0";
            sumOfCategory.EFK = "0";
            sumOfCategory.DETE = "0";
            sumOfCategory.FPA = "0";
            sumOfCategory.EpidPOT = "0";
            sumOfCategory.EpidKOT = "0";
            sumOfCategory.EpidKOT50 = "0";
            sumOfCategory.XR_1 = "0";

            return sumOfCategory;
        }

        private void importSourceBtn_Click(object sender, EventArgs e)
        {

            try {
                importSourceFile.ShowDialog();
                sourceFileLabelName.Text = importSourceFile.SafeFileName;
                sourceFileLabelPath.Text = importSourceFile.FileName;
                string fileExtension = Path.GetExtension(sourceFileLabelPath.Text);

                importFileProgressLabel.Text = "Παρακαλώ Περιμένετε..";
                importFileProgressLabel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#cf1319");
                importFileProgressLabel.Refresh();

                OleDbConnection insertConnection = new OleDbConnection();
                DataTable importedDatatable = new DataTable();

                if (fileExtension == ".xls")
                    insertConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourceFileLabelPath.Text + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                else if (fileExtension == ".xlsx")
                    insertConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFileLabelPath.Text + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                else insertConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFileLabelPath.Text + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";

                using (OleDbCommand selectAllCommand = new OleDbCommand())
                {
                    selectAllCommand.CommandText = "Select * from [Φύλλο1$]";
                    selectAllCommand.Connection = insertConnection;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = selectAllCommand;
                        da.Fill(importedDatatable);
                    }
                }

                List<ImportedData> importedDataList = new List<ImportedData>();
                ImportedData importedData = new ImportedData();

                for (int i = 0; i < importedDatatable.Rows.Count; i++) { 

                    importedData.Minas = importedDatatable.Rows[i].ItemArray[0].ToString();
                    importedData.Kata = importedDatatable.Rows[i].ItemArray[1].ToString();
                    importedData.Sys = importedDatatable.Rows[i].ItemArray[2].ToString();
                    importedData.Domi = importedDatatable.Rows[i].ItemArray[3].ToString();
                    importedData.Xrisi = importedDatatable.Rows[i].ItemArray[4].ToString();
                    importedData.Timologio = importedDatatable.Rows[i].ItemArray[5].ToString();
                    importedData.ArLogar = importedDatatable.Rows[i].ItemArray[6].ToString();
                    importedData.KWH_H = importedDatatable.Rows[i].ItemArray[7].ToString();
                    importedData.KWH_N = importedDatatable.Rows[i].ItemArray[8].ToString();
                    importedData.KWH_T = importedDatatable.Rows[i].ItemArray[9].ToString();
                    importedData.Pagio = importedDatatable.Rows[i].ItemArray[10].ToString();
                    importedData.ElXreosi = importedDatatable.Rows[i].ItemArray[11].ToString();
                    importedData.Energia = importedDatatable.Rows[i].ItemArray[12].ToString();
                    importedData.Isxys = importedDatatable.Rows[i].ItemArray[13].ToString();
                    importedData.DEnergias = importedDatatable.Rows[i].ItemArray[14].ToString();
                    importedData.CO2 = importedDatatable.Rows[i].ItemArray[15].ToString();
                    importedData.EkptEt = importedDatatable.Rows[i].ItemArray[16].ToString();
                    importedData.EkptOgk = importedDatatable.Rows[i].ItemArray[17].ToString();
                    importedData.AntEkpt = importedDatatable.Rows[i].ItemArray[18].ToString();
                    importedData.EkptTY = importedDatatable.Rows[i].ItemArray[19].ToString();
                    importedData.TK_KY = importedDatatable.Rows[i].ItemArray[20].ToString();
                    importedData.Mig = importedDatatable.Rows[i].ItemArray[21].ToString();
                    importedData.EkptPag = importedDatatable.Rows[i].ItemArray[22].ToString();
                    importedData.NeaEkptPagiou = importedDatatable.Rows[i].ItemArray[23].ToString();
                    importedData.EkptPoso = importedDatatable.Rows[i].ItemArray[24].ToString();
                    importedData.Ekpt8 = importedDatatable.Rows[i].ItemArray[25].ToString();
                    importedData.EkptosiGP = importedDatatable.Rows[i].ItemArray[26].ToString();
                    importedData.Ekpt10 = importedDatatable.Rows[i].ItemArray[27].ToString();
                    importedData.Ekpt15 = importedDatatable.Rows[i].ItemArray[28].ToString();
                    importedData.Pro = importedDatatable.Rows[i].ItemArray[29].ToString();
                    importedData.EkpOgkDim = importedDatatable.Rows[i].ItemArray[30].ToString();
                    importedData.Ekpt2 = importedDatatable.Rows[i].ItemArray[31].ToString();
                    importedData.EkptEnXT = importedDatatable.Rows[i].ItemArray[32].ToString();
                    importedData.Kouponi = importedDatatable.Rows[i].ItemArray[33].ToString();
                    importedData.EkptEbill = importedDatatable.Rows[i].ItemArray[34].ToString();
                    importedData.GreenPass = importedDatatable.Rows[i].ItemArray[35].ToString();
                    importedData.EkGreenPass = importedDatatable.Rows[i].ItemArray[36].ToString();
                    importedData.XXS = importedDatatable.Rows[i].ItemArray[37].ToString();
                    importedData.XXD = importedDatatable.Rows[i].ItemArray[38].ToString();
                    importedData.YKO = importedDatatable.Rows[i].ItemArray[39].ToString();
                    importedData.LX = importedDatatable.Rows[i].ItemArray[40].ToString();
                    importedData.ETMEAR = importedDatatable.Rows[i].ItemArray[41].ToString();
                    importedData.EFK = importedDatatable.Rows[i].ItemArray[42].ToString();
                    importedData.DETE = importedDatatable.Rows[i].ItemArray[43].ToString();
                    importedData.FPA = importedDatatable.Rows[i].ItemArray[44].ToString();
                    importedData.EpidPOT = importedDatatable.Rows[i].ItemArray[45].ToString();
                    importedData.EpidPOT = importedDatatable.Rows[i].ItemArray[46].ToString();
                    importedData.EpidKOT50 = importedDatatable.Rows[i].ItemArray[47].ToString();
                    importedData.XR_1 = importedDatatable.Rows[i].ItemArray[48].ToString();


                    importedDataList.Add(importedData);
                    importedData = new ImportedData();
                }



                for (int i = 0; i < diktiaList.Length; i++) {

                    for (int j = 0; j < XriseisList.Length; j++) {

                        ImportedData sumOfCategory = new ImportedData();
                        sumOfCategory = nullifyDataSum(sumOfCategory);
                        sumOfCategory.Sys = diktiaList[i];
                        sumOfCategory.Xrisi = XriseisList[j];


                        for (int z = 0; z < importedDataList.Count; z++)
                        {

                            if (diktiaList[i].Equals(importedDataList[z].Sys) && XriseisList[j].Equals(importedDataList[z].Xrisi)) {

                                if (importedDataList[z].KWH_H != "" && importedDataList[z].KWH_H != null)
                                    sumOfCategory.KWH_H = (double.Parse(sumOfCategory.KWH_H) + double.Parse(importedDataList[z].KWH_H)).ToString();

                                if (importedDataList[z].KWH_N != "" && importedDataList[z].KWH_N != null)
                                    sumOfCategory.KWH_N = (double.Parse(sumOfCategory.KWH_N) + double.Parse(importedDataList[z].KWH_N)).ToString();

                                if (importedDataList[z].KWH_T != "" && importedDataList[z].KWH_T != null)
                                    sumOfCategory.KWH_T = (double.Parse(sumOfCategory.KWH_T) + double.Parse(importedDataList[z].KWH_T)).ToString();

                                if (importedDataList[z].Pagio != "" && importedDataList[z].Pagio != null)
                                    sumOfCategory.Pagio = (double.Parse(sumOfCategory.Pagio) + double.Parse(importedDataList[z].Pagio)).ToString();

                                if (importedDataList[z].ElXreosi != "" && importedDataList[z].ElXreosi != null)
                                    sumOfCategory.ElXreosi = (double.Parse(sumOfCategory.ElXreosi) + double.Parse(importedDataList[z].ElXreosi)).ToString();

                                if (importedDataList[z].Energia != "" && importedDataList[z].Energia != null)
                                    sumOfCategory.Energia = (double.Parse(sumOfCategory.Energia) + double.Parse(importedDataList[z].Energia)).ToString();

                                if (importedDataList[z].Isxys != "" && importedDataList[z].Isxys != null)
                                    sumOfCategory.Isxys = (double.Parse(sumOfCategory.Isxys) + double.Parse(importedDataList[z].Isxys)).ToString();

                                if (importedDataList[z].DEnergias != "" && importedDataList[z].DEnergias != null)
                                    sumOfCategory.DEnergias = (double.Parse(sumOfCategory.DEnergias) + double.Parse(importedDataList[z].DEnergias)).ToString();
                              
                                if (importedDataList[z].CO2 != "" && importedDataList[z].CO2 != null)
                                    sumOfCategory.CO2 = (double.Parse(sumOfCategory.CO2) + double.Parse(importedDataList[z].CO2)).ToString();

                                if (importedDataList[z].EkptEt != "" && importedDataList[z].EkptEt != null)
                                    sumOfCategory.EkptEt = (double.Parse(sumOfCategory.EkptEt) + double.Parse(importedDataList[z].EkptEt)).ToString();

                                if (importedDataList[z].EkptOgk != "" && importedDataList[z].EkptOgk != null)
                                    sumOfCategory.EkptOgk = (double.Parse(sumOfCategory.EkptOgk) + double.Parse(importedDataList[z].EkptOgk)).ToString();

                                if (importedDataList[z].AntEkpt != "" && importedDataList[z].AntEkpt != null)
                                    sumOfCategory.AntEkpt = (double.Parse(sumOfCategory.AntEkpt) + double.Parse(importedDataList[z].AntEkpt)).ToString();

                                if (importedDataList[z].EkptTY != "" && importedDataList[z].EkptTY != null)
                                    sumOfCategory.EkptTY = (double.Parse(sumOfCategory.EkptTY) + double.Parse(importedDataList[z].EkptTY)).ToString();

                                if (importedDataList[z].TK_KY != "" && importedDataList[z].TK_KY != null)
                                    sumOfCategory.TK_KY = (double.Parse(sumOfCategory.TK_KY) + double.Parse(importedDataList[z].TK_KY)).ToString();

                                if (importedDataList[z].Mig != "" && importedDataList[z].Mig != null)
                                    sumOfCategory.Mig = (double.Parse(sumOfCategory.Mig) + double.Parse(importedDataList[z].Mig)).ToString();

                                if (importedDataList[z].EkptPag != "" && importedDataList[z].EkptPag != null)
                                    sumOfCategory.EkptPag = (double.Parse(sumOfCategory.EkptPag) + double.Parse(importedDataList[z].EkptPag)).ToString();

                                if (importedDataList[z].NeaEkptPagiou != "" && importedDataList[z].NeaEkptPagiou != null)
                                    sumOfCategory.NeaEkptPagiou = (double.Parse(sumOfCategory.NeaEkptPagiou) + double.Parse(importedDataList[z].NeaEkptPagiou)).ToString();

                                if (importedDataList[z].EkptPoso != "" && importedDataList[z].EkptPoso != null)
                                    sumOfCategory.EkptPoso = (double.Parse(sumOfCategory.EkptPoso) + double.Parse(importedDataList[z].EkptPoso)).ToString();

                                if (importedDataList[z].Ekpt8 != "" && importedDataList[z].Ekpt8 != null)
                                    sumOfCategory.Ekpt8 = (double.Parse(sumOfCategory.Ekpt8) + double.Parse(importedDataList[z].Ekpt8)).ToString();

                                if (importedDataList[z].EkptosiGP != "" && importedDataList[z].EkptosiGP != null)
                                    sumOfCategory.EkptosiGP = (double.Parse(sumOfCategory.EkptosiGP) + double.Parse(importedDataList[z].EkptosiGP)).ToString();

                                if (importedDataList[z].Ekpt10 != "" && importedDataList[z].Ekpt10 != null)
                                    sumOfCategory.Ekpt10 = (double.Parse(sumOfCategory.Ekpt10) + double.Parse(importedDataList[z].Ekpt10)).ToString();

                                if (importedDataList[z].Ekpt15 != "" && importedDataList[z].Ekpt15 != null)
                                    sumOfCategory.Ekpt15 = (double.Parse(sumOfCategory.Ekpt15) + double.Parse(importedDataList[z].Ekpt15)).ToString();

                                if (importedDataList[z].Pro != "" && importedDataList[z].Pro != null)
                                    sumOfCategory.Pro = (double.Parse(sumOfCategory.Pro) + double.Parse(importedDataList[z].Pro)).ToString();

                                if (importedDataList[z].EkpOgkDim != "" && importedDataList[z].EkpOgkDim != null)
                                    sumOfCategory.EkpOgkDim = (double.Parse(sumOfCategory.EkpOgkDim) + double.Parse(importedDataList[z].EkpOgkDim)).ToString();

                                if (importedDataList[z].Ekpt2 != "" && importedDataList[z].Ekpt2 != null)
                                    sumOfCategory.Ekpt2 = (double.Parse(sumOfCategory.Ekpt2) + double.Parse(importedDataList[z].Ekpt2)).ToString();

                                if (importedDataList[z].EkptEnXT != "" && importedDataList[z].EkptEnXT != null)
                                    sumOfCategory.EkptEnXT = (double.Parse(sumOfCategory.EkptEnXT) + double.Parse(importedDataList[z].EkptEnXT)).ToString();

                                if (importedDataList[z].Kouponi != "" && importedDataList[z].Kouponi != null)
                                    sumOfCategory.Kouponi = (double.Parse(sumOfCategory.Kouponi) + double.Parse(importedDataList[z].Kouponi)).ToString();

                                if (importedDataList[z].EkptEbill != "" && importedDataList[z].EkptEbill != null)
                                    sumOfCategory.EkptEbill = (double.Parse(sumOfCategory.EkptEbill) + double.Parse(importedDataList[z].EkptEbill)).ToString();

                                if (importedDataList[z].GreenPass != "" && importedDataList[z].GreenPass != null)
                                    sumOfCategory.GreenPass = (double.Parse(sumOfCategory.GreenPass) + double.Parse(importedDataList[z].GreenPass)).ToString();

                                if (importedDataList[z].EkGreenPass != "" && importedDataList[z].EkGreenPass != null)
                                    sumOfCategory.EkGreenPass = (double.Parse(sumOfCategory.EkGreenPass) + double.Parse(importedDataList[z].EkGreenPass)).ToString();

                                if (importedDataList[z].XXS != "" && importedDataList[z].XXS != null)
                                    sumOfCategory.XXS = (double.Parse(sumOfCategory.XXS) + double.Parse(importedDataList[z].XXS)).ToString();

                                if (importedDataList[z].XXD != "" && importedDataList[z].XXD != null)
                                    sumOfCategory.XXD = (double.Parse(sumOfCategory.XXD) + double.Parse(importedDataList[z].XXD)).ToString();

                                if (importedDataList[z].YKO != "" && importedDataList[z].YKO != null)
                                    sumOfCategory.YKO = (double.Parse(sumOfCategory.YKO) + double.Parse(importedDataList[z].YKO)).ToString();

                                if (importedDataList[z].LX != "" && importedDataList[z].LX != null)
                                    sumOfCategory.LX = (double.Parse(sumOfCategory.LX) + double.Parse(importedDataList[z].LX)).ToString();

                                if (importedDataList[z].ETMEAR != "" && importedDataList[z].ETMEAR != null)
                                    sumOfCategory.ETMEAR = (double.Parse(sumOfCategory.ETMEAR) + double.Parse(importedDataList[z].ETMEAR)).ToString();

                                if (importedDataList[z].EFK != "" && importedDataList[z].EFK != null)
                                    sumOfCategory.EFK = (double.Parse(sumOfCategory.EFK) + double.Parse(importedDataList[z].EFK)).ToString();

                                if (importedDataList[z].DETE != "" && importedDataList[z].DETE != null)
                                    sumOfCategory.DETE = (double.Parse(sumOfCategory.DETE) + double.Parse(importedDataList[z].DETE)).ToString();

                                if (importedDataList[z].FPA != "" && importedDataList[z].FPA != null)
                                    sumOfCategory.FPA = (double.Parse(sumOfCategory.FPA) + double.Parse(importedDataList[z].FPA)).ToString();

                                if (importedDataList[z].EpidPOT != "" && importedDataList[z].EpidPOT != null)
                                    sumOfCategory.EpidPOT = (double.Parse(sumOfCategory.EpidPOT) + double.Parse(importedDataList[z].EpidPOT)).ToString();

                                if (importedDataList[z].EpidKOT != "" && importedDataList[z].EpidKOT != null)
                                    sumOfCategory.EpidKOT = (double.Parse(sumOfCategory.EpidKOT) + double.Parse(importedDataList[z].EpidKOT)).ToString();

                                if (importedDataList[z].EpidKOT50 != "" && importedDataList[z].EpidKOT50 != null)
                                    sumOfCategory.EpidKOT50 = (double.Parse(sumOfCategory.EpidKOT50) + double.Parse(importedDataList[z].EpidKOT50)).ToString();

                                if (importedDataList[z].XR_1 != "" && importedDataList[z].XR_1 != null)
                                    sumOfCategory.KWH_H = (double.Parse(sumOfCategory.XR_1) + double.Parse(importedDataList[z].XR_1)).ToString();

                            }
                        }

                        if (XriseisList[j].IndexOf("ΠΚΥ", StringComparison.OrdinalIgnoreCase) >= 0) {
                            dataSumPKYList.Add(sumOfCategory.ImportedDataClone());
                        }
                        else {
                            dataSumDEIList.Add(sumOfCategory.ImportedDataClone());
                        }

                    }

                }

                importFileProgressLabel.Text = "Επιτυχής Φόρτωση!";
                importFileProgressLabel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#008080");
                importFileProgressLabel.Refresh();

                ExportBtnDEI.Enabled = true;
                ExportBtnPKY.Enabled = true;

            }
            catch (Exception exc) {
                importFileProgressLabel.Text = "Σφάλμα..";
                importFileProgressLabel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#cf1319");
                importFileProgressLabel.Refresh();
                MessageBox.Show(exc.Message);
            }
        }

        private void ExportBtn_Click(object sender, EventArgs e)
        {
            exportFileSourceDEI.ShowDialog();
            exportFileNameLabelDEI.Text = exportFileSourceDEI.SafeFileName;
            exportFilePathLabelDEI.Text = exportFileSourceDEI.FileName; 
            string fileExtension = Path.GetExtension(sourceFileLabelPath.Text);
            goDEIbtn.Enabled = true;

        }


        private void goDEIbtn_Click(object sender, EventArgs e)
        {
            exportFileProgressLabelDEI.Text = "Παρακαλώ Περιμένετε..";
            exportFileProgressLabelDEI.ForeColor = System.Drawing.ColorTranslator.FromHtml("#cf1319");
            exportFileProgressLabelDEI.Refresh();

            OleDbConnection exportConnection = new OleDbConnection();
            DataTable exportedDatatable = new DataTable();

            try
            {
                // Open the document for editing.
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(exportFilePathLabelDEI.Text, true))
                {
                    // Access the main Workbook part, which contains all references.
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                    // get sheet by name
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "ΑΡΧΙΚΑ").FirstOrDefault();
                    // get worksheetpart by sheet id
                    WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    int month = monthSelectbox.SelectedIndex;
                    int column = month * 4 - 1;
                    int reset = 0;
                    int i = column;

                    for (int j = 0; j < dataSumDEIList.Count; j++)
                    {

                        reset++;
                        if (reset == 8)
                        {
                            reset = 0;
                            i++;
                        }

                        if (dataSumDEIList[j].Xrisi == "ΑΓΡΟΤΙΚΗ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "D", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumDEIList[j].KWH_H);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "L", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumDEIList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AB", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumDEIList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AJ", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumDEIList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AR", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumDEIList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "AZ", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumDEIList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BH", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumDEIList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BP", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumDEIList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "BX", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumDEIList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CF", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumDEIList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CN", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumDEIList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CV", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumDEIList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DD", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumDEIList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DL", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumDEIList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DT", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumDEIList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EB", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumDEIList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EJ", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumDEIList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "ER", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumDEIList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "EZ", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumDEIList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FH", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumDEIList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FP", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumDEIList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "FX", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumDEIList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GF", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumDEIList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GN", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumDEIList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GV", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumDEIList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HD", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumDEIList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HL", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumDEIList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HT", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumDEIList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IB", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumDEIList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IJ", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumDEIList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IR", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumDEIList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "IZ", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumDEIList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JH", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumDEIList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JP", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumDEIList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "JX", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumDEIList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KF", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumDEIList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);

                        }
                        else if (dataSumDEIList[j].Xrisi == "ΒΙΟΜΗΧΑΝΙΚΗ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "E", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumDEIList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "M", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumDEIList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AC", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumDEIList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AK", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumDEIList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AS", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumDEIList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BA", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumDEIList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BI", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumDEIList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BQ", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumDEIList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "BY", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumDEIList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CG", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumDEIList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CO", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumDEIList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CW", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumDEIList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DE", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumDEIList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DM", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumDEIList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DU", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumDEIList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EC", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumDEIList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EK", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumDEIList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "ES", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumDEIList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FA", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumDEIList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FI", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumDEIList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FQ", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumDEIList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "FY", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumDEIList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GG", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumDEIList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GO", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumDEIList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GW", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumDEIList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HE", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumDEIList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HM", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumDEIList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HU", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumDEIList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IC", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumDEIList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IK", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumDEIList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IS", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumDEIList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JA", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumDEIList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JI", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumDEIList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JQ", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumDEIList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "JY", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumDEIList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KG", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumDEIList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumDEIList[j].Xrisi == "ΓΕΝΙΚΗ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "F", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumDEIList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "N", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumDEIList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AD", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumDEIList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AL", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumDEIList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AT", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumDEIList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BB", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumDEIList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BJ", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumDEIList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BR", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumDEIList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "BZ", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumDEIList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CH", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumDEIList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CP", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumDEIList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CX", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumDEIList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DF", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumDEIList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DN", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumDEIList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DV", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumDEIList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "ED", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumDEIList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EL", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumDEIList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "ET", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumDEIList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FB", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumDEIList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FJ", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumDEIList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FR", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumDEIList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "FZ", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumDEIList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GH", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumDEIList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GP", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumDEIList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GX", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumDEIList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HF", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumDEIList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HN", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumDEIList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HV", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumDEIList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "ID", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumDEIList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IL", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumDEIList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IT", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumDEIList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JB", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumDEIList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JJ", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumDEIList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JR", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumDEIList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "JZ", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumDEIList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KH", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumDEIList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumDEIList[j].Xrisi == "ΔΗΜΟΣΙΑ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "G", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumDEIList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "M", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumDEIList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AE", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumDEIList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AM", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumDEIList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AU", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumDEIList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BC", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumDEIList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BK", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumDEIList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BR", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumDEIList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CA", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumDEIList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CI", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumDEIList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CQ", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumDEIList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CY", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumDEIList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DG", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumDEIList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DO", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumDEIList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DW", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumDEIList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EE", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumDEIList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EM", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumDEIList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EU", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumDEIList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FC", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumDEIList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FK", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumDEIList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FS", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumDEIList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GA", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumDEIList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GI", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumDEIList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GQ", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumDEIList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GY", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumDEIList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HG", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumDEIList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HO", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumDEIList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HW", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumDEIList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IE", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumDEIList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IM", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumDEIList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IU", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumDEIList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JC", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumDEIList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JK", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumDEIList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JS", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumDEIList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KA", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumDEIList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KI", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumDEIList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumDEIList[j].Xrisi == "ΚΟΙΝΟΤΙΚΑ - ΦΟΠ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "H", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumDEIList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "N", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumDEIList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AF", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumDEIList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AN", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumDEIList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AV", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumDEIList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BD", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumDEIList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BL", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumDEIList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BS", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumDEIList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CB", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumDEIList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CJ", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumDEIList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CR", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumDEIList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CZ", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumDEIList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DH", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumDEIList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DP", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumDEIList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DX", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumDEIList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EF", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumDEIList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EN", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumDEIList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EV", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumDEIList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FD", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumDEIList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FL", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumDEIList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FT", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumDEIList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GB", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumDEIList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GJ", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumDEIList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GR", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumDEIList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GZ", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumDEIList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HH", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumDEIList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HP", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumDEIList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HX", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumDEIList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IF", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumDEIList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IN", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumDEIList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IV", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumDEIList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JD", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumDEIList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JL", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumDEIList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JT", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumDEIList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KB", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumDEIList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KJ", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumDEIList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumDEIList[j].Xrisi == "ΝΠΔΔ - Δ. ΕΠΙΧ. - ΠΡ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "I", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumDEIList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "M", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumDEIList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AG", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumDEIList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AO", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumDEIList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AW", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumDEIList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BE", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumDEIList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BM", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumDEIList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BT", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumDEIList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CC", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumDEIList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CK", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumDEIList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CS", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumDEIList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "DA", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumDEIList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DI", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumDEIList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DQ", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumDEIList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DY", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumDEIList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EG", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumDEIList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EO", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumDEIList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EW", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumDEIList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FE", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumDEIList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FM", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumDEIList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FU", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumDEIList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GC", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumDEIList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GK", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumDEIList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GS", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumDEIList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "HA", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumDEIList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HI", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumDEIList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HQ", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumDEIList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HY", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumDEIList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IG", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumDEIList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IO", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumDEIList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IW", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumDEIList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JE", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumDEIList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JM", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumDEIList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JU", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumDEIList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KC", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumDEIList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KK", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumDEIList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumDEIList[j].Xrisi == "ΟΙΚΙΑΚΗ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "J", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumDEIList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "N", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumDEIList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AH", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumDEIList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AP", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumDEIList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AX", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumDEIList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BF", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumDEIList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BN", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumDEIList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BU", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumDEIList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CD", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumDEIList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CL", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumDEIList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CT", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumDEIList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "DB", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumDEIList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DJ", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumDEIList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DR", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumDEIList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DZ", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumDEIList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EH", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumDEIList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EP", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumDEIList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EX", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumDEIList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FF", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumDEIList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FN", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumDEIList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FV", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumDEIList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GD", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumDEIList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GL", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumDEIList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GT", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumDEIList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "HB", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumDEIList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HJ", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumDEIList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HR", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumDEIList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HZ", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumDEIList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IH", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumDEIList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IP", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumDEIList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IX", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumDEIList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JF", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumDEIList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JN", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumDEIList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JV", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumDEIList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KD", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumDEIList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KL", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumDEIList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }

                        // Save the worksheet.
                        worksheetPart.Worksheet.Save();
                        // for recacluation of formula
                        spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                        spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                    }

                }

                exportFileProgressLabelDEI.Text = "Επιτυχία!";
                exportFileProgressLabelDEI.ForeColor = System.Drawing.ColorTranslator.FromHtml("#008080");
                exportFileProgressLabelDEI.Refresh();

            }
            catch (Exception exc)
            {
                exportFileProgressLabelDEI.Text = "Σφάλμα..";
                exportFileProgressLabelDEI.ForeColor = System.Drawing.ColorTranslator.FromHtml("#cf1319");
                exportFileProgressLabelDEI.Refresh();
                MessageBox.Show(exc.Message);
            }
        }

        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null) return null;

            var FirstRow = row.Elements<Cell>().Where(c => string.Compare
            (c.CellReference.Value, columnName +
            rowIndex, true) == 0).FirstOrDefault();

            if (FirstRow == null) return null;

            return FirstRow;
        }

        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            Row row = worksheet.GetFirstChild<SheetData>().
            Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                throw new ArgumentException(String.Format("No row with index {0} found in spreadsheet", rowIndex));
            }
            return row;
        }

        private void ExportBtnPKY_Click(object sender, EventArgs e)
        {
            exportFileSourcePKY.ShowDialog();
            exportFileNameLabelPKY.Text = exportFileSourcePKY.SafeFileName;
            exportFilePathLabelPKY.Text = exportFileSourcePKY.FileName;
            string fileExtension = Path.GetExtension(sourceFileLabelPath.Text);
            goPKYbtn.Enabled = true;
        }

        private void goPKYbtn_Click(object sender, EventArgs e)
        {
            exportFileProgressLabelPKY.Text = "Παρακαλώ Περιμένετε..";
            exportFileProgressLabelPKY.ForeColor = System.Drawing.ColorTranslator.FromHtml("#cf1319");
            exportFileProgressLabelPKY.Refresh();

            OleDbConnection exportConnection = new OleDbConnection();
            DataTable exportedDatatable = new DataTable();

            try
            {
                // Open the document for editing.
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(exportFilePathLabelPKY.Text, true))
                {
                    // Access the main Workbook part, which contains all references.
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                    // get sheet by name
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "ΑΡΧΙΚΑ").FirstOrDefault();
                    // get worksheetpart by sheet id
                    WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    int month = monthSelectbox.SelectedIndex;
                    int column = month * 4 - 1;
                    int reset = 0;
                    int i = column;

                    for (int j = 0; j < dataSumPKYList.Count; j++)
                    {

                        reset++;
                        if (reset == 8)
                        {
                            reset = 0;
                            i++;
                        }

                        if (dataSumPKYList[j].Xrisi == "ΑΓΡΟΤΙΚΗ ΠΚΥ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "D", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumPKYList[j].KWH_H);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "L", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumPKYList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AB", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumPKYList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AJ", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumPKYList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AR", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumPKYList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "AZ", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumPKYList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BH", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumPKYList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BP", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumPKYList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "BX", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumPKYList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CF", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumPKYList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CN", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumPKYList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CV", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumPKYList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DD", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumPKYList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DL", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumPKYList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DT", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumPKYList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EB", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumPKYList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EJ", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumPKYList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "ER", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumPKYList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "EZ", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumPKYList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FH", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumPKYList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FP", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumPKYList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "FX", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumPKYList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GF", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumPKYList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GN", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumPKYList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GV", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumPKYList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HD", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumPKYList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HL", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumPKYList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HT", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumPKYList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IB", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumPKYList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IJ", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumPKYList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IR", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumPKYList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "IZ", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumPKYList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JH", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumPKYList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JP", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumPKYList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "JX", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumPKYList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KF", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumPKYList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);

                        }
                        else if (dataSumPKYList[j].Xrisi == "ΒΙΟΜΗΧΑΝΙΚΗ ΠΚΥ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "E", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumPKYList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "M", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumPKYList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AC", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumPKYList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AK", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumPKYList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AS", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumPKYList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BA", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumPKYList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BI", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumPKYList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BQ", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumPKYList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "BY", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumPKYList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CG", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumPKYList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CO", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumPKYList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CW", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumPKYList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DE", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumPKYList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DM", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumPKYList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DU", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumPKYList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EC", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumPKYList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EK", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumPKYList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "ES", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumPKYList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FA", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumPKYList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FI", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumPKYList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FQ", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumPKYList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "FY", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumPKYList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GG", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumPKYList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GO", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumPKYList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GW", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumPKYList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HE", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumPKYList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HM", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumPKYList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HU", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumPKYList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IC", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumPKYList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IK", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumPKYList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IS", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumPKYList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JA", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumPKYList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JI", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumPKYList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JQ", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumPKYList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "JY", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumPKYList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KG", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumPKYList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumPKYList[j].Xrisi == "ΓΕΝΙΚΗ ΠΚΥ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "F", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumPKYList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "N", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumPKYList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AD", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumPKYList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AL", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumPKYList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AT", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumPKYList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BB", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumPKYList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BJ", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumPKYList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BR", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumPKYList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "BZ", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumPKYList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CH", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumPKYList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CP", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumPKYList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CX", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumPKYList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DF", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumPKYList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DN", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumPKYList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DV", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumPKYList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "ED", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumPKYList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EL", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumPKYList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "ET", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumPKYList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FB", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumPKYList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FJ", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumPKYList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FR", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumPKYList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "FZ", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumPKYList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GH", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumPKYList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GP", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumPKYList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GX", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumPKYList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HF", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumPKYList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HN", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumPKYList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HV", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumPKYList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "ID", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumPKYList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IL", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumPKYList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IT", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumPKYList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JB", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumPKYList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JJ", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumPKYList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JR", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumPKYList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "JZ", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumPKYList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KH", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumPKYList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumPKYList[j].Xrisi == "ΔΗΜΟΣΙΑ ΠΚΥ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "G", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumPKYList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "M", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumPKYList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AE", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumPKYList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AM", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumPKYList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AU", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumPKYList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BC", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumPKYList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BK", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumPKYList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BR", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumPKYList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CA", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumPKYList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CI", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumPKYList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CQ", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumPKYList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CY", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumPKYList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DG", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumPKYList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DO", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumPKYList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DW", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumPKYList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EE", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumPKYList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EM", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumPKYList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EU", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumPKYList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FC", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumPKYList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FK", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumPKYList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FS", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumPKYList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GA", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumPKYList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GI", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumPKYList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GQ", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumPKYList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GY", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumPKYList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HG", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumPKYList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HO", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumPKYList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HW", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumPKYList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IE", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumPKYList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IM", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumPKYList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IU", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumPKYList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JC", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumPKYList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JK", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumPKYList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JS", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumPKYList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KA", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumPKYList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KI", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumPKYList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumPKYList[j].Xrisi == "ΚΟΙΝΟΤΙΚΑ - ΦΟΠ ΠΚΥ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "H", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumPKYList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "N", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumPKYList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AF", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumPKYList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AN", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumPKYList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AV", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumPKYList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BD", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumPKYList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BL", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumPKYList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BS", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumPKYList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CB", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumPKYList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CJ", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumPKYList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CR", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumPKYList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "CZ", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumPKYList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DH", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumPKYList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DP", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumPKYList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DX", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumPKYList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EF", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumPKYList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EN", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumPKYList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EV", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumPKYList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FD", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumPKYList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FL", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumPKYList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FT", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumPKYList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GB", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumPKYList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GJ", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumPKYList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GR", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumPKYList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "GZ", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumPKYList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HH", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumPKYList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HP", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumPKYList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HX", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumPKYList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IF", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumPKYList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IN", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumPKYList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IV", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumPKYList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JD", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumPKYList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JL", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumPKYList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JT", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumPKYList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KB", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumPKYList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KJ", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumPKYList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumPKYList[j].Xrisi == "ΝΠΔΔ - Δ. ΕΠΙΧ. - ΠΡ ΠΚΥ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "I", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumPKYList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "M", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumPKYList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AG", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumPKYList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AO", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumPKYList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AW", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumPKYList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BE", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumPKYList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BM", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumPKYList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BT", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumPKYList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CC", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumPKYList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CK", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumPKYList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CS", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumPKYList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "DA", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumPKYList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DI", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumPKYList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DQ", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumPKYList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DY", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumPKYList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EG", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumPKYList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EO", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumPKYList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EW", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumPKYList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FE", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumPKYList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FM", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumPKYList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FU", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumPKYList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GC", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumPKYList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GK", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumPKYList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GS", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumPKYList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "HA", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumPKYList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HI", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumPKYList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HQ", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumPKYList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HY", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumPKYList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IG", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumPKYList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IO", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumPKYList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IW", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumPKYList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JE", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumPKYList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JM", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumPKYList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JU", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumPKYList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KC", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumPKYList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KK", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumPKYList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }
                        else if (dataSumPKYList[j].Xrisi == "ΟΙΚΙΑΚΗ ΠΚΥ")
                        {
                            Cell cellKWH_H = GetCell(worksheetPart.Worksheet, "J", (uint)i);
                            cellKWH_H.CellValue = new CellValue(dataSumPKYList[j].KWH_H);
                            cellKWH_H.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKWH_N = GetCell(worksheetPart.Worksheet, "N", (uint)i);
                            cellKWH_N.CellValue = new CellValue(dataSumPKYList[j].KWH_N);
                            cellKWH_N.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPagio = GetCell(worksheetPart.Worksheet, "AH", (uint)i);
                            cellPagio.CellValue = new CellValue(dataSumPKYList[j].Pagio);
                            cellPagio.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellElXreosi = GetCell(worksheetPart.Worksheet, "AP", (uint)i);
                            cellElXreosi.CellValue = new CellValue(dataSumPKYList[j].ElXreosi);
                            cellElXreosi.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEnergia = GetCell(worksheetPart.Worksheet, "AX", (uint)i);
                            cellEnergia.CellValue = new CellValue(dataSumPKYList[j].Energia);
                            cellEnergia.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellIsxys = GetCell(worksheetPart.Worksheet, "BF", (uint)i);
                            cellIsxys.CellValue = new CellValue(dataSumPKYList[j].Isxys);
                            cellIsxys.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEt = GetCell(worksheetPart.Worksheet, "BN", (uint)i);
                            cellEkptEt.CellValue = new CellValue(dataSumPKYList[j].EkptEt);
                            cellEkptEt.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptOgk = GetCell(worksheetPart.Worksheet, "BU", (uint)i);
                            cellEkptOgk.CellValue = new CellValue(dataSumPKYList[j].EkptOgk);
                            cellEkptOgk.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptTY = GetCell(worksheetPart.Worksheet, "CD", (uint)i);
                            cellEkptTY.CellValue = new CellValue(dataSumPKYList[j].EkptTY);
                            cellEkptTY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellTK_KY = GetCell(worksheetPart.Worksheet, "CL", (uint)i);
                            cellTK_KY.CellValue = new CellValue(dataSumPKYList[j].TK_KY);
                            cellTK_KY.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellMig = GetCell(worksheetPart.Worksheet, "CT", (uint)i);
                            cellMig.CellValue = new CellValue(dataSumPKYList[j].Mig);
                            cellMig.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPag = GetCell(worksheetPart.Worksheet, "DB", (uint)i);
                            cellEkptPag.CellValue = new CellValue(dataSumPKYList[j].EkptPag);
                            cellEkptPag.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt10 = GetCell(worksheetPart.Worksheet, "DJ", (uint)i);
                            cellEkpt10.CellValue = new CellValue(dataSumPKYList[j].Ekpt10);
                            cellEkpt10.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt15 = GetCell(worksheetPart.Worksheet, "DR", (uint)i);
                            cellEkpt15.CellValue = new CellValue(dataSumPKYList[j].Ekpt15);
                            cellEkpt15.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellPro = GetCell(worksheetPart.Worksheet, "DZ", (uint)i);
                            cellPro.CellValue = new CellValue(dataSumPKYList[j].Pro);
                            cellPro.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpOgkDim = GetCell(worksheetPart.Worksheet, "EH", (uint)i);
                            cellEkpOgkDim.CellValue = new CellValue(dataSumPKYList[j].EkpOgkDim);
                            cellEkpOgkDim.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt8 = GetCell(worksheetPart.Worksheet, "EP", (uint)i);
                            cellEkpt8.CellValue = new CellValue(dataSumPKYList[j].Ekpt8);
                            cellEkpt8.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptosiGP = GetCell(worksheetPart.Worksheet, "EX", (uint)i);
                            cellEkptosiGP.CellValue = new CellValue(dataSumPKYList[j].EkptosiGP);
                            cellEkptosiGP.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXS = GetCell(worksheetPart.Worksheet, "FF", (uint)i);
                            cellXXS.CellValue = new CellValue(dataSumPKYList[j].XXS);
                            cellXXS.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXXD = GetCell(worksheetPart.Worksheet, "FN", (uint)i);
                            cellXXD.CellValue = new CellValue(dataSumPKYList[j].XXD);
                            cellXXD.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellLX = GetCell(worksheetPart.Worksheet, "FV", (uint)i);
                            cellLX.CellValue = new CellValue(dataSumPKYList[j].LX);
                            cellLX.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellYKO = GetCell(worksheetPart.Worksheet, "GD", (uint)i);
                            cellYKO.CellValue = new CellValue(dataSumPKYList[j].YKO);
                            cellYKO.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellETMEAR = GetCell(worksheetPart.Worksheet, "GL", (uint)i);
                            cellETMEAR.CellValue = new CellValue(dataSumPKYList[j].ETMEAR);
                            cellETMEAR.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEFK = GetCell(worksheetPart.Worksheet, "GT", (uint)i);
                            cellEFK.CellValue = new CellValue(dataSumPKYList[j].EFK);
                            cellEFK.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellDETE = GetCell(worksheetPart.Worksheet, "HB", (uint)i);
                            cellDETE.CellValue = new CellValue(dataSumPKYList[j].DETE);
                            cellDETE.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellFPA = GetCell(worksheetPart.Worksheet, "HJ", (uint)i);
                            cellFPA.CellValue = new CellValue(dataSumPKYList[j].FPA);
                            cellFPA.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidPOT = GetCell(worksheetPart.Worksheet, "HR", (uint)i);
                            cellEpidPOT.CellValue = new CellValue(dataSumPKYList[j].EpidPOT);
                            cellEpidPOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT = GetCell(worksheetPart.Worksheet, "HZ", (uint)i);
                            cellEpidKOT.CellValue = new CellValue(dataSumPKYList[j].EpidKOT);
                            cellEpidKOT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEpidKOT50 = GetCell(worksheetPart.Worksheet, "IH", (uint)i);
                            cellEpidKOT50.CellValue = new CellValue(dataSumPKYList[j].EpidKOT50);
                            cellEpidKOT50.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellXR_1 = GetCell(worksheetPart.Worksheet, "IP", (uint)i);
                            cellXR_1.CellValue = new CellValue(dataSumPKYList[j].XR_1);
                            cellXR_1.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellCO2 = GetCell(worksheetPart.Worksheet, "IX", (uint)i);
                            cellCO2.CellValue = new CellValue(dataSumPKYList[j].CO2);
                            cellCO2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellNeaEkptPagiou = GetCell(worksheetPart.Worksheet, "JF", (uint)i);
                            cellNeaEkptPagiou.CellValue = new CellValue(dataSumPKYList[j].NeaEkptPagiou);
                            cellNeaEkptPagiou.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptPoso = GetCell(worksheetPart.Worksheet, "JN", (uint)i);
                            cellEkptPoso.CellValue = new CellValue(dataSumPKYList[j].EkptPoso);
                            cellEkptPoso.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkpt2 = GetCell(worksheetPart.Worksheet, "JV", (uint)i);
                            cellEkpt2.CellValue = new CellValue(dataSumPKYList[j].Ekpt2);
                            cellEkpt2.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellEkptEnXT = GetCell(worksheetPart.Worksheet, "KD", (uint)i);
                            cellEkptEnXT.CellValue = new CellValue(dataSumPKYList[j].EkptEnXT);
                            cellEkptEnXT.DataType = new EnumValue<CellValues>(CellValues.Number);

                            Cell cellKouponi = GetCell(worksheetPart.Worksheet, "KL", (uint)i);
                            cellKouponi.CellValue = new CellValue(dataSumPKYList[j].Kouponi);
                            cellKouponi.DataType = new EnumValue<CellValues>(CellValues.Number);
                        }

                        // Save the worksheet.
                        worksheetPart.Worksheet.Save();
                        // for recacluation of formula
                        spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                        spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                    }

                }

                exportFileProgressLabelPKY.Text = "Επιτυχία!";
                exportFileProgressLabelPKY.ForeColor = System.Drawing.ColorTranslator.FromHtml("#008080");
                exportFileProgressLabelPKY.Refresh();
            }
            catch (Exception exc)
            {
                exportFileProgressLabelPKY.Text = "Σφάλμα..";
                exportFileProgressLabelPKY.ForeColor = System.Drawing.ColorTranslator.FromHtml("#cf1319");
                exportFileProgressLabelPKY.Refresh();
                MessageBox.Show(exc.Message);
            }
        }
    }
}
