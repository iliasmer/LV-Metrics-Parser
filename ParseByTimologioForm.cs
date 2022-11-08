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
    public partial class ParseByTimologioForm : Form
    {
        Dictionary<string, string> timologioMapDEI = new Dictionary<string, string>();
        Dictionary<string, string> timologioMapPKY = new Dictionary<string, string>();
        Dictionary<string, string> timologioMapAll = new Dictionary<string, string>();
        List<string> distinctTimologioDEI = new List<string>();
        List<string> distinctTimologioPKY = new List<string>();
        List<ImportedData> dataSumDEIList = new List<ImportedData>();
        List<ImportedData> dataSumPKYList = new List<ImportedData>();
        string[] diktiaList = new string[] { "1", "2", "3", "4" };

        public ParseByTimologioForm()
        {
            InitializeComponent();
            initializeTimologioMapDEI();
            initializeTimologioMapPKY();
            initializeTimologioMapAll();
        }

        private void ParseByTimologioForm_Load(object sender, EventArgs e)
        {
            sourceFileLabelName.Text = "";
            importFileProgressLabel.Text = "";
            exportFileNameLabelDEI.Text = "";
            exportFileProgressLabelDEI.Text = "";
            exportFileNameLabelPKY.Text = "";
            exportFileProgressLabelPKY.Text = "";
        }

        void getDistinctTimologioListPKY(List<ImportedData> importedDataList)
        {

            for (int i = 0; i < importedDataList.Count; i++)
            {
                if (!existsInList(distinctTimologioPKY, timologioMapAll[importedDataList[i].Timologio]) && importedDataList[i].Timologio.IndexOf("ΠΚΥ", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    distinctTimologioPKY.Add(timologioMapAll[importedDataList[i].Timologio]);
                }
            }

        }

        void getDistinctTimologioListDEI(List<ImportedData> importedDataList) {

            for (int i = 0; i < importedDataList.Count; i++) {
                if (!existsInList(distinctTimologioDEI, timologioMapAll[importedDataList[i].Timologio]) && importedDataList[i].Timologio.IndexOf("ΠΚΥ", StringComparison.OrdinalIgnoreCase) < 0) {
                    distinctTimologioDEI.Add(timologioMapAll[importedDataList[i].Timologio]);
                }
            }

        }

        void normalizePKY(List<ImportedData> importedDataList) { 
            for (int i = 0; i < importedDataList.Count; i++)
            {
                importedDataList[i].Timologio = importedDataList[i].Timologio.Replace("ΠKY", "ΠΚΥ").Replace("ΠΚY", "ΠΚΥ").Replace("ΠKΥ", "ΠΚΥ");
            }
        }

        bool existsInList(List<string> distinctTimologio, string val) {
            bool f = false;
            for (int i = 0; i < distinctTimologio.Count; i++) {
                if (distinctTimologio[i].Equals(val)) {
                    f = true;
                    break;
                }
            }
            return f;
        }


        void initializeTimologioMapAll()
        {
            timologioMapAll.Add("Γ1~Οικιακό Τιμολόγιο~χωρίς χρονοχρ 1Φ", "Γ1 1Φ");
            timologioMapAll.Add("ΓΤ~Οικιακό Τιμολόγιο~4 Παιδιά 1Φ", "Γ1 1Φ");
            timologioMapAll.Add("Μη αντιστ.", "Γ1 1Φ");
            timologioMapAll.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ1Φ~KOTA", "Γ1 1Φ ΚΑ");
            timologioMapAll.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ1Φ~KOTB", "Γ1 1Φ ΚΒ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Βορράς 1Φ", "ΓΠ 1Φ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Κεντρ Ελ 1Φ", "ΓΠ 1Φ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Νότος 1Φ", "ΓΠ 1Φ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Μονοφασικό", "ΓΠ 1Φ");
            timologioMapAll.Add("Γ1~Οικιακό Τιμολόγιο~χωρίς χρονοχρ 3Φ", "Γ1 3Φ");
            timologioMapAll.Add("ΓΤ~Οικιακό Τιμολόγιο~5-9 Παιδιά 1Φ", "Γ1 3Φ");
            timologioMapAll.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ3Φ~KOTA", "Γ1 3Φ ΚΑ");
            timologioMapAll.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ3Φ~KOTB", "Γ1 3Φ ΚΒ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Νότος 3Φ", "ΓΠ 3Φ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Βορράς 3Φ", "ΓΠ 3Φ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Κεντρ Ελ 3Φ", "ΓΠ 3Φ");
            timologioMapAll.Add("ΓΠ~Οικιακό Τιμολόγιο~Τριφασικό", "ΓΠ 3Φ");
            timologioMapAll.Add("Γ1Ν~Οικιακό Τιμολόγιο~με χρονοχρέωση 1Φ", "Γ1Ν 1Φ");
            timologioMapAll.Add("ΓΤΝ~Οικιακό Τιμολόγιο~4 Παιδιά 1Φ χρον", "Γ1Ν 1Φ");
            timologioMapAll.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση1Φ~KOTA", "Γ1Ν 1Φ ΚΑ");
            timologioMapAll.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση1Φ~KOTB", "Γ1Ν 1Φ ΚΒ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Βορράς 1Φ χρονοχ", "ΓΠΝ 1Φ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Κεντρ Ελ 1Φ χρον", "ΓΠΝ 1Φ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Νότος 1Φ χρονοχρ", "ΓΠΝ 1Φ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~ 1Φ χρονοχ", "ΓΠΝ 1Φ");
            timologioMapAll.Add("Γ1Ν~Οικιακό Τιμολόγιο~με χρονοχρέωση 3Φ", "Γ1Ν 3Φ");
            timologioMapAll.Add("ΓΤΝ~Οικιακό Τιμολόγιο~4 Παιδιά 3Φ χρον", "Γ1Ν 3Φ");
            timologioMapAll.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση3Φ~KOTA", "Γ1Ν 3Φ ΚΑ");
            timologioMapAll.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση3Φ~KOTB", "Γ1Ν 3Φ ΚΒ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Νότος 3Φ χρονοχρ", "ΓΠΝ 3Φ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Βορράς 3Φ χρονοχ", "ΓΠΝ 3Φ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Κεντρ Ελ 3Φ χρον", "ΓΠΝ 3Φ");
            timologioMapAll.Add("ΓΠΝ~Οικιακό Τιμολόγιο~ 3Φ χρονοχ", "ΓΠΝ 3Φ");
            timologioMapAll.Add("Γ21~Επαγγελματικό~Εμπορική χρήση", "Γ21");
            timologioMapAll.Add("Γ21~Επαγγελματ.ΠΚΥ(Α)~ Αγροτική Χρήσ", "Γ21");
            timologioMapAll.Add("Γ22~Επαγγελματικό~με ισχύ Εμπορ χρ", "Γ22 Ι");
            timologioMapAll.Add("Γ22~Επαγγελματικό~με άεργα Εμπορ χρ", "Γ22 Α");
            timologioMapAll.Add("Γ23~Επαγγελματικό~με χρον Εμπορ χρ", "Γ23");
            timologioMapAll.Add("Γ21~Επαγγελματικό~Βιομηχανική χρήση", "Γ21 Β");
            timologioMapAll.Add("Γ22~Επαγγελματικό~με άεργα Βιομηχ χρ", "Γ22 Α Β");
            timologioMapAll.Add("Γ22~Επαγγελματικό~με ισχύ Βιομηχ χρ", "Γ22 Ι Β");
            timologioMapAll.Add("Γ23~Επαγγελματικό~με χρον Βιομηχ χρ", "Γ23 Β");
            timologioMapAll.Add("Ε21~Εταιρικό~Βιομηχανική χρήση", "Γ21 ΕΤ");
            timologioMapAll.Add("Ε22~Εταιρικό~με άεργα Βιομηχ χρ", "Γ22 Α Β ΕΤ");
            timologioMapAll.Add("Ε22~Εταιρικό~με ισχύ Βιομηχ χρ", "Γ22 Ι Β ΕΤ");
            timologioMapAll.Add("Ε21~Εταιρικό~Εμπορική χρήση", "Γ21 ΕΤ");
            timologioMapAll.Add("Ε22~Εταιρικό~με ισχύ Εμπορ χρ", "Γ22 Ι ΕΤ");
            timologioMapAll.Add("Ε22~Εταιρικό~με άεργα Εμπορ χρ", "Γ22 Α ΕΤ");
            timologioMapAll.Add("Ε23~Εταιρικό~με χρον Εμπορ χρ", "Γ23 ΕΤ");
            timologioMapAll.Add("~Αγροτικό Τιμ/γιο~", "ΜΑΤ");
            timologioMapAll.Add("ΦΟΠ~Τιμολόγιο", "ΦΟΠ");
            timologioMapAll.Add("Γ1~myHome Online~χωρίς χρονο~Digital", "myHome Online χωρίς χρονο");
            timologioMapAll.Add("Γ1~myHome Online~χωρχρο~KOTA~Digital", "myHome Online χωρχρο KOTA");
            timologioMapAll.Add("Γ1~myHome Online~χωρχρο~KOTB~Digital", "myHome Online χωρχρο KOTB");
            timologioMapAll.Add("Γ1Ν~myHome Online~χρονο~KOTA~Digital", "myHome Online χρονο KOTA");
            timologioMapAll.Add("Γ1Ν~myHome Online~χρονο~KOTB~Digital", "myHome Online χρονο KOTΒ");
            timologioMapAll.Add("Γ1Ν~myHome Online~χρονοχρέωσ~Digital", "myHome Online με χρονο");
            timologioMapAll.Add("ΓΠ~myHome Online~Digital", "myHome Online ΓΠ");
            timologioMapAll.Add("ΓΠΝ~myHome Online~χρονοχ~Digital", "myHome Online χρονοχ ΓΠ");
            timologioMapAll.Add("~myHome Enter N~με χρονοχρέωση~HB", "myHome Enter με χρονοχρέωση");
            timologioMapAll.Add("~myHome Enter ΓΠ~Μονοφασικό~HB", "myHome Enter ΓΠ");
            timologioMapAll.Add("~myHome Enter ΓΠΝ~χρονοχ~HB", "myHome Enter ΓΠ χρονοχ");
            timologioMapAll.Add("~myHome Enter~χωρίς χρονοχρ~HB", "myHome Enterχωρίς χρονοχρ");
            timologioMapAll.Add("ΠΚΥ~Επαγ/τικό Χρον.~Εμπορική χρ", "Γ23 ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Επαγγελματικό~Εμπορική χρήση", "Γ21 ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό Χρονοχρ.~1Φ", "Γ1Ν 1Φ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό Χρονοχρ.~1Φ~KOTA", "Γ1Ν 1Φ ΚΑ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό Χρονοχρ.~1Φ~KOTB", "Γ1Ν 1Φ ΚΒ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό Χρονοχρ.~3Φ", "Γ1Ν 3Φ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό Χρονοχρ.~3Φ~KOTA", "Γ1Ν 3Φ ΚΑ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό Χρονοχρ.~3Φ~KOTB", "Γ1Ν 3Φ ΚΒ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό~χωρίς χρονοχρ 1Φ", "Γ1 1Φ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό~χωρίς χρονοχρ 3Φ", "Γ1 3Φ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό~χωχρονοχρ1Φ~KOTA", "Γ1 1Φ ΚΑ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό~χωχρονοχρ1Φ~KOTB", "Γ1 1Φ ΚΒ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό~χωχρονοχρ3Φ~KOTA", "Γ1 3Φ ΚΑ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Οικιακό~χωχρονοχρ3Φ~KOTB", "Γ1 3Φ ΚΒ ΠΚΥ");
            timologioMapAll.Add("ΠΚΥ~Αγροτικό~Αγροτική χρήση", "ΜΑΤ ΠΚΥ");
        }


        void initializeTimologioMapPKY() {
            timologioMapPKY.Add("ΠΚΥ~Επαγ/τικό Χρον.~Εμπορική χρ", "Γ23 ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Επαγγελματικό~Εμπορική χρήση", "Γ21 ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό Χρονοχρ.~1Φ", "Γ1Ν 1Φ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό Χρονοχρ.~1Φ~KOTA", "Γ1Ν 1Φ ΚΑ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό Χρονοχρ.~1Φ~KOTB", "Γ1Ν 1Φ ΚΒ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό Χρονοχρ.~3Φ", "Γ1Ν 3Φ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό Χρονοχρ.~3Φ~KOTA", "Γ1Ν 3Φ ΚΑ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό Χρονοχρ.~3Φ~KOTB", "Γ1Ν 3Φ ΚΒ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό~χωρίς χρονοχρ 1Φ", "Γ1 1Φ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό~χωρίς χρονοχρ 3Φ", "Γ1 3Φ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό~χωχρονοχρ1Φ~KOTA", "Γ1 1Φ ΚΑ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό~χωχρονοχρ1Φ~KOTB", "Γ1 1Φ ΚΒ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό~χωχρονοχρ3Φ~KOTA", "Γ1 3Φ ΚΑ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Οικιακό~χωχρονοχρ3Φ~KOTB", "Γ1 3Φ ΚΒ ΠΚΥ");
            timologioMapPKY.Add("ΠΚΥ~Αγροτικό~Αγροτική χρήση", "ΜΑΤ ΠΚΥ");
        }

        void initializeTimologioMapDEI() {
            timologioMapDEI.Add("Γ1~Οικιακό Τιμολόγιο~χωρίς χρονοχρ 1Φ", "Γ1 1Φ");
            timologioMapDEI.Add("ΓΤ~Οικιακό Τιμολόγιο~4 Παιδιά 1Φ", "Γ1 1Φ");
            timologioMapDEI.Add("Μη αντιστ.", "Γ1 1Φ");
            timologioMapDEI.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ1Φ~KOTA", "Γ1 1Φ ΚΑ");
            timologioMapDEI.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ1Φ~KOTB", "Γ1 1Φ ΚΒ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Βορράς 1Φ", "ΓΠ 1Φ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Κεντρ Ελ 1Φ", "ΓΠ 1Φ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Νότος 1Φ", "ΓΠ 1Φ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Μονοφασικό", "ΓΠ 1Φ");
            timologioMapDEI.Add("Γ1~Οικιακό Τιμολόγιο~χωρίς χρονοχρ 3Φ", "Γ1 3Φ");
            timologioMapDEI.Add("ΓΤ~Οικιακό Τιμολόγιο~5-9 Παιδιά 1Φ", "Γ1 3Φ");
            timologioMapDEI.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ3Φ~KOTA", "Γ1 3Φ ΚΑ");
            timologioMapDEI.Add("Γ1~Οικιακό Τιμολόγιο~χωρίςχρονοχρ3Φ~KOTB", "Γ1 3Φ ΚΒ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Νότος 3Φ", "ΓΠ 3Φ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Βορράς 3Φ", "ΓΠ 3Φ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Κεντρ Ελ 3Φ", "ΓΠ 3Φ");
            timologioMapDEI.Add("ΓΠ~Οικιακό Τιμολόγιο~Τριφασικό", "ΓΠ 3Φ");
            timologioMapDEI.Add("Γ1Ν~Οικιακό Τιμολόγιο~με χρονοχρέωση 1Φ", "Γ1Ν 1Φ");
            timologioMapDEI.Add("ΓΤΝ~Οικιακό Τιμολόγιο~4 Παιδιά 1Φ χρον", "Γ1Ν 1Φ");
            timologioMapDEI.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση1Φ~KOTA", "Γ1Ν 1Φ ΚΑ");
            timologioMapDEI.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση1Φ~KOTB", "Γ1Ν 1Φ ΚΒ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Βορράς 1Φ χρονοχ", "ΓΠΝ 1Φ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Κεντρ Ελ 1Φ χρον", "ΓΠΝ 1Φ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Νότος 1Φ χρονοχρ", "ΓΠΝ 1Φ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~ 1Φ χρονοχ", "ΓΠΝ 1Φ");
            timologioMapDEI.Add("Γ1Ν~Οικιακό Τιμολόγιο~με χρονοχρέωση 3Φ", "Γ1Ν 3Φ");
            timologioMapDEI.Add("ΓΤΝ~Οικιακό Τιμολόγιο~4 Παιδιά 3Φ χρον", "Γ1Ν 3Φ");
            timologioMapDEI.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση3Φ~KOTA", "Γ1Ν 3Φ ΚΑ");
            timologioMapDEI.Add("Γ1Ν~Οικιακό Τιμολόγιο~χρονοχρέωση3Φ~KOTB", "Γ1Ν 3Φ ΚΒ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Νότος 3Φ χρονοχρ", "ΓΠΝ 3Φ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Βορράς 3Φ χρονοχ", "ΓΠΝ 3Φ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~Κεντρ Ελ 3Φ χρον", "ΓΠΝ 3Φ");
            timologioMapDEI.Add("ΓΠΝ~Οικιακό Τιμολόγιο~ 3Φ χρονοχ", "ΓΠΝ 3Φ");
            timologioMapDEI.Add("Γ21~Επαγγελματικό~Εμπορική χρήση", "Γ21");
            timologioMapDEI.Add("Γ21~Επαγγελματ.ΠΚΥ(Α)~ Αγροτική Χρήσ", "Γ21");
            timologioMapDEI.Add("Γ22~Επαγγελματικό~με ισχύ Εμπορ χρ", "Γ22 Ι");
            timologioMapDEI.Add("Γ22~Επαγγελματικό~με άεργα Εμπορ χρ", "Γ22 Α");
            timologioMapDEI.Add("Γ23~Επαγγελματικό~με χρον Εμπορ χρ", "Γ23");
            timologioMapDEI.Add("Γ21~Επαγγελματικό~Βιομηχανική χρήση", "Γ21 Β");
            timologioMapDEI.Add("Γ22~Επαγγελματικό~με άεργα Βιομηχ χρ", "Γ22 Α Β");
            timologioMapDEI.Add("Γ22~Επαγγελματικό~με ισχύ Βιομηχ χρ", "Γ22 Ι Β");
            timologioMapDEI.Add("Γ23~Επαγγελματικό~με χρον Βιομηχ χρ", "Γ23 Β");
            timologioMapDEI.Add("Ε21~Εταιρικό~Βιομηχανική χρήση", "Γ21 ΕΤ");
            timologioMapDEI.Add("Ε22~Εταιρικό~με άεργα Βιομηχ χρ", "Γ22 Α Β ΕΤ");
            timologioMapDEI.Add("Ε22~Εταιρικό~με ισχύ Βιομηχ χρ", "Γ22 Ι Β ΕΤ");
            timologioMapDEI.Add("Ε21~Εταιρικό~Εμπορική χρήση", "Γ21 ΕΤ");
            timologioMapDEI.Add("Ε22~Εταιρικό~με ισχύ Εμπορ χρ", "Γ22 Ι ΕΤ");
            timologioMapDEI.Add("Ε22~Εταιρικό~με άεργα Εμπορ χρ", "Γ22 Α ΕΤ");
            timologioMapDEI.Add("Ε23~Εταιρικό~με χρον Εμπορ χρ", "Γ23 ΕΤ");
            timologioMapDEI.Add("~Αγροτικό Τιμ/γιο~", "ΜΑΤ");
            timologioMapDEI.Add("ΦΟΠ~Τιμολόγιο", "ΦΟΠ");
            timologioMapDEI.Add("Γ1~myHome Online~χωρίς χρονο~Digital", "myHome Online χωρίς χρονο");
            timologioMapDEI.Add("Γ1~myHome Online~χωρχρο~KOTA~Digital", "myHome Online χωρχρο KOTA");
            timologioMapDEI.Add("Γ1~myHome Online~χωρχρο~KOTB~Digital", "myHome Online χωρχρο KOTB");
            timologioMapDEI.Add("Γ1Ν~myHome Online~χρονο~KOTA~Digital", "myHome Online χρονο KOTA");
            timologioMapDEI.Add("Γ1Ν~myHome Online~χρονο~KOTB~Digital", "myHome Online χρονο KOTΒ");
            timologioMapDEI.Add("Γ1Ν~myHome Online~χρονοχρέωσ~Digital", "myHome Online με χρονο");
            timologioMapDEI.Add("ΓΠ~myHome Online~Digital", "myHome Online ΓΠ");
            timologioMapDEI.Add("ΓΠΝ~myHome Online~χρονοχ~Digital", "myHome Online χρονοχ ΓΠ");
            timologioMapDEI.Add("~myHome Enter N~με χρονοχρέωση~HB", "myHome Enter με χρονοχρέωση");
            timologioMapDEI.Add("~myHome Enter ΓΠ~Μονοφασικό~HB", "myHome Enter ΓΠ");
            timologioMapDEI.Add("~myHome Enter ΓΠΝ~χρονοχ~HB", "myHome Enter ΓΠ χρονοχ");
            timologioMapDEI.Add("~myHome Enter~χωρίς χρονοχρ~HB", "myHome Enterχωρίς χρονοχρ");
        }

        private ImportedData nullifyDataSum(ImportedData sumOfCategory)
        {
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

                for (int i = 0; i < importedDatatable.Rows.Count; i++)
                {

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

                distinctTimologioDEI = new List<string>();
                distinctTimologioPKY = new List<string>();
                normalizePKY(importedDataList);
                getDistinctTimologioListDEI(importedDataList);
                getDistinctTimologioListPKY(importedDataList);

                //Creation of DEI list
                for (int i = 0; i < diktiaList.Length; i++) {

                    for (int j = 0; j < distinctTimologioDEI.Count; j++)
                    {
                        ImportedData sumOfCategory = new ImportedData();
                        sumOfCategory = nullifyDataSum(sumOfCategory);
                        sumOfCategory.Sys = diktiaList[i];

                        for (int z = 0; z < importedDataList.Count; z++)
                        {
                            if (diktiaList[i].Equals(importedDataList[z].Sys) && distinctTimologioDEI[j].Equals(timologioMapAll[importedDataList[z].Timologio]))
                            {
                                if (importedDataList[z].Timologio != "" && importedDataList[z].Timologio != null)
                                    sumOfCategory.Timologio = timologioMapDEI[importedDataList[z].Timologio];

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

                        dataSumDEIList.Add(sumOfCategory.ImportedDataClone());
                    }

                }


                //Creation of PKY list
                for (int i = 0; i < diktiaList.Length; i++)
                {

                    for (int j = 0; j < distinctTimologioPKY.Count; j++)
                    {
                        ImportedData sumOfCategory = new ImportedData();
                        sumOfCategory = nullifyDataSum(sumOfCategory);
                        sumOfCategory.Sys = diktiaList[i];

                        for (int z = 0; z < importedDataList.Count; z++)
                        {
                            if (diktiaList[i].Equals(importedDataList[z].Sys) && distinctTimologioPKY[j].Equals(timologioMapAll[importedDataList[z].Timologio]))
                            {
                                if (importedDataList[z].Timologio != "" && importedDataList[z].Timologio != null)
                                    sumOfCategory.Timologio = timologioMapPKY[importedDataList[z].Timologio];

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

                        dataSumPKYList.Add(sumOfCategory.ImportedDataClone());
                    }

                }


                importFileProgressLabel.Text = "Επιτυχής Φόρτωση!";
                importFileProgressLabel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#008080");
                importFileProgressLabel.Refresh();

                ExportBtnDEI.Enabled = true;

            }
            catch (Exception exc) {
                importFileProgressLabel.Text = "Σφάλμα..";
                importFileProgressLabel.ForeColor = System.Drawing.ColorTranslator.FromHtml("#cf1319");
                importFileProgressLabel.Refresh();
                MessageBox.Show(exc.Message);
            }
        }

    }
}
