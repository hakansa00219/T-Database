using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

namespace TencereDB
{
    public partial class FormStartup : Form
    {
        //dragging
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HTCAPTION = 0x2;
        [DllImport("User32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        //

        NewTencere form1 = new NewTencere();
        static public List<Tencere> TencereList = new List<Tencere>();
        static public int selectedIndex = -1;
        public FormStartup()
        { 
            InitializeComponent();
            tabControl1.TabPages.Remove(tabPage3);
            tabPage1.BackColor = Color.White;
            //tabControl1.DrawMode = TabDrawMode.OwnerDrawFixed;
            //tabControl1.DrawItem += tabControl1_DrawItem;
            menuStrip1.Items[0].ForeColor = Color.Red;
            menuStrip1.Items[1].ForeColor = Color.Red;
           // tabControl1.DrawMode = TabDrawMode.Normal;
        }
        /*
        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index == tabControl1.SelectedIndex)
            {
                e.Graphics.DrawString(tabControl1.TabPages[e.Index].Text,
                    new Font(tabControl1.Font, FontStyle.Bold),
                    Brushes.Red,
                    new PointF(e.Bounds.X + 3, e.Bounds.Y + 3));
            }
            else
            {
                e.Graphics.DrawString(tabControl1.TabPages[e.Index].Text,
                    tabControl1.Font,
                    Brushes.DarkRed,
                    new PointF(e.Bounds.X + 3, e.Bounds.Y + 3));
            }
        }*/
        AppSettings defaultAppSettings = new AppSettings()
        {
            DBPath = "TencereDB.json",
            JsonName = "Tencereler.json",
            ExcellName = "Tencereler.xlsx "
        };

        private void FormStartup_Load(object sender, EventArgs e)
        {
            if (File.Exists("TencereDB.json"))
            {
                // buraya bakılacak
                // read it
            }
            else
            {
                using (StreamWriter sw = File.AppendText("TencereDB.json"))
                {
                    sw.WriteLine(JsonConvert.SerializeObject(defaultAppSettings));
                    sw.Close();
                }
            }

            readAppSettings();

            if (!checkForNetwork())
                MessageBox.Show("no network");

            if (!File.Exists(defaultAppSettings.DBPath + defaultAppSettings.JsonName))
            {
                using (StreamWriter sw = File.AppendText(defaultAppSettings.DBPath + defaultAppSettings.JsonName))
                {

                    sw.Write("");
                    sw.Close();
                }
            }
            else
            {

            }
            readJsonData();
            updateListView();
            //WriteExcel();
        }

        public List<Tencere> GetList()
        {
            return TencereList;
        }
        void readAppSettings()
        {
            var lines = File.ReadAllLines("TencereDB.json");
            defaultAppSettings = JsonConvert.DeserializeObject<AppSettings>(lines[0]);
        }

        bool checkForNetwork()
        {
            try
            {
                Ping myPing = new Ping();
                String host = "google.com";
                byte[] buffer = new byte[32];
                int timeout = 1000;
                PingOptions pingOptions = new PingOptions();
                PingReply reply = myPing.Send(host, timeout, buffer, pingOptions);
                return (reply.Status == IPStatus.Success);
            }
            catch (Exception)
            {
                return false;
            }
        }
        public void editTencere(string[] specs, int[] Dims)
        {
            TencereList[selectedIndex].ID = specs[0];
            TencereList[selectedIndex].Tip = specs[1];
            TencereList[selectedIndex].Uretici = specs[2];
            TencereList[selectedIndex].DimT = Dims[0];
            TencereList[selectedIndex].DimN = Dims[1];

            if (TencereList[selectedIndex].ID == String.Empty)
            {
                TencereList[selectedIndex].ID = "unknown";
            }
            if (TencereList[selectedIndex].Tip == String.Empty)
            {
                TencereList[selectedIndex].Tip = "unknown";
            }
            if (TencereList[selectedIndex].Uretici == String.Empty)
            {
                TencereList[selectedIndex].Uretici = "unknown";
            }
            updateListView();
            updateJson();

        }

        void readJsonData()
        {
            TencereList.Clear();
            string[] lines = File.ReadAllLines(defaultAppSettings.DBPath + defaultAppSettings.JsonName);
            if (lines.Length != 0)
            {
                foreach (string line in lines)
                {
                    Tencere tenc = JsonConvert.DeserializeObject<Tencere>(line);
                    TencereList.Add(tenc);
                }
            }
        }
        public static void colorListViewHeader(ref ListView list, Color backColor, Color foreColor)
        {
            list.OwnerDraw = true;
            list.DrawColumnHeader +=
                new DrawListViewColumnHeaderEventHandler
                (
                    (sender, e) => headerDraw(sender, e, backColor, foreColor)
                );
            list.DrawItem += new DrawListViewItemEventHandler(bodyDraw);
        }

        private static void headerDraw(object sender, DrawListViewColumnHeaderEventArgs e, Color backColor, Color foreColor)
        {
            using (SolidBrush backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
            }

            using (SolidBrush foreBrush = new SolidBrush(foreColor))
            {
                e.Graphics.DrawString(e.Header.Text, e.Font, foreBrush, e.Bounds);
            }
        }

        private static void bodyDraw(object sender, DrawListViewItemEventArgs e)
        {
            e.DrawDefault = true;
        }
        void InitListView()
        {
            listView1.Columns.Add("ID", 30);
            listView1.BackColor = Color.White;
            listView1.ForeColor = Color.Red;
            listView1.Columns.Add("Tip", 50);
            listView1.Columns.Add("Üretici", 50);
            listView1.Columns.Add("Boyut", 50);
            colorListViewHeader(ref listView1, Color.Red, Color.White);
            listView1.Columns[0].ListView.Font = new Font(listView1.Columns[0].ListView.Font, FontStyle.Bold);
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.AutoArrange = true;

        }
        public void updateJson()
        {
            File.Create(defaultAppSettings.DBPath + defaultAppSettings.JsonName).Close();
            using (StreamWriter sw = File.AppendText(defaultAppSettings.DBPath + defaultAppSettings.JsonName))
            {

                foreach (Tencere item in TencereList)
                {
                    sw.WriteLine(JsonConvert.SerializeObject(item));
                }
                sw.Close();
            }
        }
        void updateListView()
        {
            listView1.Clear();
            InitListView();
            for (int a = 0; a < TencereList.Count; a++)
            {
                string[] arr = new string[5];
                ListViewItem itm;
                arr[0] = TencereList[a].ID;
                arr[1] = TencereList[a].Tip;
                arr[2] = TencereList[a].Uretici;
                arr[3] = TencereList[a].DimT.ToString();
                arr[4] = TencereList[a].DimN.ToString();
                itm = new ListViewItem(arr);
                itm.Font = new Font("TimesNewRoman",9);
              
                listView1.Items.Add(itm);

            }
            for (int a1 = 0; a1 < 4; a1++)
            {
                listView1.Columns[a1].Width = -2;
            }

        }
        private void TrackBar1_Scroll(object sender, EventArgs e)
        {
            textBox1.Text = trackBar1.Value.ToString();
            
        }
        private void MakeVisible()
        {
            label36.Visible = true;
            label37.Visible = true;
            label38.Visible = true;
            label39.Visible = true;
            label42.Visible = true;
            label43.Visible = true;
            label44.Visible = true;
            label45.Visible = true;
            label46.Visible = true;
        }
        private void TrackBar1_ValueChanged(object sender, EventArgs e)
        {
            MakeVisible();
            string frequency = textBox1.Text;
            int _frequency = Convert.ToInt32(frequency);
            try
            {
                for (int i = 0; i < 4; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {

                        _lbl[i, j].Size = new Size(100, 23);
                        _lbl[i, j].Location = new Point((j * 125) + 20, (i * 41) + 20);
                        _lbl[i, j].Text = "--";
                        _lbl[i, j].Font = new Font("Arial", 11);
                        try
                        {
                            switch (i)
                            {

                                case 0:
                                    _lbl[i, j].Text = string.Format("{0:0.####}", TencereList[index].specsClass[j][_frequency - 10].R);
                                    break;
                                case 1:
                                    _lbl[i, j].Text = string.Format("{0:0.#####E+00}", TencereList[index].specsClass[j][_frequency - 10].L);
                                    break;
                                case 2:
                                    _lbl[i, j].Text = string.Format("{0:0.####}", TencereList[index].specsClass[j][_frequency - 10].Z);
                                    break;
                                case 3:
                                    _lbl[i, j].Text = string.Format("{0:0.####E+00}", TencereList[index].specsClass[j][_frequency - 10].C);
                                    break;
                            }


                        }
                        catch (Exception)
                        {


                        }
                        panel3.Controls.Add(_lbl[i, j]);
                        if (_lbl[i, j].Text == string.Empty)
                        {
                            _lbl[i, j].Text = "--";
                        }

                    }
                }
                panel3.Show();
            } catch
            {

            }
        }
        private void MakeNonVisible()
        {
            label36.Visible = false;
            label37.Visible = false;
            label38.Visible = false;
            label39.Visible = false;
            label42.Visible = false;
            label43.Visible = false;
            label44.Visible = false;
            label45.Visible = false;
            label46.Visible = false;
        }
        Label[,] _lbl = new Label[4, 5];
        int index;
        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            MakeNonVisible();
            if (listView1.SelectedIndices.Count <= 0)
                return;
            panel3.Controls.Clear();
            index = listView1.SelectedItems[0].Index;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
            label1.Text = TencereList[index].ID;
            label2.Text = TencereList[index].Tip;
            label3.Text = TencereList[index].Uretici;
            label4.Text = TencereList[index].DimT.ToString() + "-" + TencereList[index].DimN.ToString();
            for(int i = 0; i < 4; i++)
            {
                for(int j = 0; j < 5; j++)
                {
                    _lbl[i, j] = new Label();
                }
            }

            selectedIndex = index;
        }
        public static string txt1, txt2, txt3, txt4, txt5;
        private void Button1_Click(object sender, EventArgs e)
        {
            // yeni
            //panel2.Visible = true;
            txt1 = textBox2.Text;
            txt2 = textBox3.Text;
            txt3 = textBox4.Text;
            txt4 = customTextBox1.Text;
            txt5 = customTextBox2.Text;
            NewTencere form1 = new NewTencere();
            form1.Show();

        }
        public string[,,] checkTextboxes()
        {
            string[,,] changedSpecs = new string[5, 91, 4];
            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 91; j++)
                {
                    for (int k = 0; k < 4; k++)
                    {
                        string tbName = ("T" + i.ToString() + j.ToString() + k.ToString());
                        changedSpecs[i, j, k] = panel2.Controls[tbName].Text;
                    }
                }
            }

            return changedSpecs;
        }

        private void Button5_Click(object sender, EventArgs e)
        {

            string[] s2 = new string[3];
            int[] dims = new int[2];
            s2[0] = textBox2.Text;
            s2[1] = textBox3.Text;
            s2[2] = textBox4.Text;
            string d = customTextBox1.Text;
            string d2 = customTextBox2.Text;
            dims[0] = Convert.ToInt32(d);
            dims[1] = Convert.ToInt32(d2);
            editTencere(s2, dims);
            tabControl1.SelectedTab = tabControl1.TabPages[0];
            panel2.Visible = false;
            MessageBox.Show("Tencere belirtilen şekilde değiştirildi.");
        }

        private void listView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (listView1.FocusedItem.Bounds.Contains(e.Location))
                {
                    contextMenuStrip1.Show(MousePosition);

                }
            }
        }
        public int indx;
        private void DeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedIndices.Count > 0)
            {
                foreach (ListViewItem item in listView1.SelectedItems)
                {
                    TencereList.RemoveAt(item.Index);
                    indx = item.Index;
                    item.Remove();
                }
            }
        }
        private void YenileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                readJsonData();
                updateListView();
                MessageBox.Show("Liste yenilendi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Problem çıktı.");
            }
        }

        private void KaydetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                updateJson();
                readJsonData();
                updateListView();
                MessageBox.Show("Kaydedildi.");
            }catch(Exception)
            {
                MessageBox.Show("Problem çıktı.");
            }
            
        }
        private void TextBox35_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        public void ToChangeValues()
        {
            textBox2.Text = TencereList[selectedIndex].ID;
            textBox3.Text = TencereList[selectedIndex].Tip;
            textBox4.Text = TencereList[selectedIndex].Uretici;
            customTextBox1.Text = TencereList[selectedIndex].DimT.ToString();
            customTextBox2.Text = TencereList[selectedIndex].DimN.ToString();
            panel2.Visible = true;

        }

        private void Button2_Click(object sender, EventArgs e)
        {

            if (selectedIndex == -1)
            {
                tabControl1.SelectedTab = tabControl1.TabPages[0];
                MessageBox.Show("Lütfen listeden değiştirmek istediğiniz tencereyi seçin.");
            }
            else
            {
                //panel2.Visible = true;
                button5.Text = "Değiştir";
                ToChangeValues();
            }
        }
        private void Button4_Click(object sender, EventArgs e)
        {
            Photo form2 = new Photo();
            form2.Show();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateJson();
            readJsonData();
            updateListView();
            if (tabControl1.SelectedTab == tabControl1.TabPages[0])
            {
                panel2.Visible = false;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[1])
            {
                if (selectedIndex == -1)
                {
                    //do nothing.

                }
                else
                {
                    try
                    {
                        if (TencereList.Count > 0)
                        {
                            label17.Text = TencereList[selectedIndex].ID;
                        }
                    }
                    catch
                    {

                    }
                    

                }
            }
            if(tabControl1.SelectedTab != tabPage3) tabControl1.TabPages.Remove(tabPage3);
        }

        private void DeğiştirToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // edit
            if (selectedIndex == -1)
                return;
            tabControl1.SelectedTab = tabControl1.TabPages[1];
            label17.Text = TencereList[selectedIndex].ID;
        }



        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                trackBar1.Value = 10;
            }
            else
            {
                if (Convert.ToInt32(textBox1.Text) < 10)
                {
                    trackBar1.Value = 10;
                }
                else if (Convert.ToInt32(textBox1.Text) > 100)
                {
                    trackBar1.Value = 100;
                }
                else
                {
                    trackBar1.Value = Convert.ToInt32(textBox1.Text);
                }
            }
        }

        private void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        string ExcelPath = Environment.CurrentDirectory + @"\DByeni.xlsx";
        public void WriteExcel()
        {
            
            if (File.Exists(ExcelPath)) 
            {
                 
                //File.Delete(@"D:\TencereDB\TencereDB\bin\Debug\DByeni.xlsx");
            }else
            {
                File.Create(ExcelPath);
            }
            //open excel
            Excel.Application exl = new Excel.Application();
            Excel.Workbook wb = exl.Workbooks.Add();//Open(ExcelPath);
            Excel.Worksheet[] ws = new Excel.Worksheet[TencereList.Count];
            exl.Visible = false;
            //for (int q = 0; q < wb.Worksheets.Count - ws.Length; q++) wb.Sheets[wb.Worksheets.Count - ws.Length].Delete();
            //for (int j = 0; j < wb.Worksheets.Count-1; j++) wb.Worksheets[2].Delete();
            if (wb.Worksheets.Count != TencereList.Count+1)
            {
                for (int i = 0; i < TencereList.Count; i++)
                {
                    //open workbook and worksheet
                    ws[i] = wb.Worksheets.Add();
                    ws[i] = wb.ActiveSheet;
                    ws[i].Name =  TencereList[i].ID;
                    //headers
                    ws[i].Cells[1, 1] = "ID";
                    ws[i].Cells[2, 1] = "Tür";
                    ws[i].Cells[3, 1] = "Üretici";
                    ws[i].Cells[4, 1] = "TabanÇap";
                    ws[i].Cells[5, 1] = "NormÇap";
                    //headers
                    ws[i].Cells[2, 3] = "Frekans";
                    ws[i].Cells[1, 4] = "145";
                    ws[i].Cells[1, 8] = "180";
                    ws[i].Cells[1, 12] = "210";
                    ws[i].Cells[1, 16] = "240";
                    ws[i].Cells[1, 20] = "270";
                    for (int j = 0; j < 5; j++)
                    {
                        ws[i].Cells[2, 4 + j * 4] = "R";
                        ws[i].Cells[2, 5 + j * 4] = "L";
                        ws[i].Cells[2, 6 + j * 4] = "Z";
                        ws[i].Cells[2, 7 + j * 4] = "C";
                    }
                    //data/
                    for(int j1 = 0; j1 < 91; j1++)
                    {
                        ws[i].Cells[j1 + 3, 3] = j1 + 10 + " kHz";
                        for(int j2 = 0; j2 < TencereList[i].specsClass.Length; j2++)
                        {
                            try
                            {
                                ws[i].Cells[j1 + 3, 4 + 4 * j2] = TencereList[i].specsClass[j2][j1].R; //?? nullchecker;
                            }
                            catch (NullReferenceException)
                            {
                                ws[i].Cells[j1 + 3, 4 + 4 * j2] = 0;
                            }
                            try
                            {
                                ws[i].Cells[j1 + 3, 4 * j2 + 5] = TencereList[i].specsClass[j2][j1].L;
                            }
                            catch (NullReferenceException)
                            {
                                ws[i].Cells[j1 + 3, 4 * j2 + 5] = 0;
                            }
                            try
                            {
                                ws[i].Cells[j1 + 3, 4 * j2 + 6] = TencereList[i].specsClass[j2][j1].Z;
                            }
                            catch (NullReferenceException)
                            {
                                ws[i].Cells[j1 + 3, 4 * j2 + 6] = 0;
                            }
                            try
                            {
                                ws[i].Cells[j1 + 3, 4 * j2 + 7] = TencereList[i].specsClass[j2][j1].C;
                            }
                            catch (NullReferenceException)
                            {
                                ws[i].Cells[j1 + 3, 4 * j2 + 7] = 0;
                            }



                        }
                        
                    }
                    ws[i].Cells[1, 2] = TencereList[i].ID;
                    ws[i].Cells[2, 2] = TencereList[i].Tip;
                    ws[i].Cells[3, 2] = TencereList[i].Uretici;
                    ws[i].Cells[4, 2] = TencereList[i].DimT;
                    ws[i].Cells[5, 2] = TencereList[i].DimN;

                    //do these for every coil in the list
                    ws[i].Range[ws[i].Cells[1, 4], ws[i].Cells[1, 7]].Merge();
                    ws[i].Range[ws[i].Cells[1, 8], ws[i].Cells[1, 11]].Merge();
                    ws[i].Range[ws[i].Cells[1, 12], ws[i].Cells[1, 15]].Merge();
                    ws[i].Range[ws[i].Cells[1, 16], ws[i].Cells[1, 19]].Merge();
                    ws[i].Range[ws[i].Cells[1, 20], ws[i].Cells[1, 23]].Merge();
                    //settings
                    ws[i].Columns[2].ColumnWidth = 18;
                    ws[i].Columns[1].ColumnWidth = 10;
                    ws[i].get_Range("B1", "B5").Font.Color = Color.Red;
                    ws[i].get_Range("C3", "C93").Font.Color = Color.Red;
                    ws[i].get_Range("C3", "C93").Font.Bold = true;

                    ws[i].get_Range("A1", "A5").Font.Bold = true;
                    ws[i].get_Range("A1", "A5").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    ws[i].get_Range("C1", "W2").Font.Bold = true;
                    ws[i].get_Range("C1", "W2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    
                }
            } 
            
            exl.DisplayAlerts = false;
            exl.UserControl = false;
            wb.SaveAs(ExcelPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close(true);
            exl.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(exl);
            MessageBox.Show("Excel dosyası kaydedildi.");
        }
        public void ReadExcel()
        {
            // listedeki herşeyi siliyoruz.
            TencereList.Clear();
            
            //open exel       
            Excel.Application exl = new Excel.Application();
            Excel.Workbook wb = exl.Workbooks.Open(ExcelPath);
            Excel.Worksheet[] ws = new Excel.Worksheet[wb.Worksheets.Count-1];
            Excel.Range range;
            for (int ii = 0; ii < wb.Worksheets.Count-1 ; ii++) ws[ii] = wb.Sheets[ii+1];
            exl.Visible = false;
            Tencere[] tencere = new Tencere[wb.Worksheets.Count - 1];
             
            // yeni tencere oluşturuyoruz her exceldeki sayfa için
            for(int ij = 0; ij < wb.Worksheets.Count - 1; ij ++)
            {
                tencere[ij] = new Tencere();
            }
           
            
            for (int i = 0; i < wb.Worksheets.Count-1; i++)
            {
                //data/
                tencere[i].specsClass = new TencereSpecs[5][];
                for (int i1 = 0; i1 < 5; i1++)
                {
                    tencere[i].specsClass[i1] = new TencereSpecs[91];
                }
                for (int j1 = 0; j1 < 91; j1++)
                {
                    for (int j2 = 0; j2 < 5; j2++)
                    {
                        tencere[i].specsClass[j2][j1] = new TencereSpecs();
                        range = ws[i].Cells[j1 + 3, 4 * j2 + 4];
                        tencere[i].specsClass[j2][j1].R = range.Value;
                        range = ws[i].Cells[j1 + 3, 4 * j2 + 5];
                        tencere[i].specsClass[j2][j1].L = range.Value;
                        range = ws[i].Cells[j1 + 3, 4 * j2 + 6];
                        tencere[i].specsClass[j2][j1].Z = range.Value;
                        range = ws[i].Cells[j1 + 3, 4 * j2 + 7];
                        tencere[i].specsClass[j2][j1].C = range.Value;
                    }

                }
                range = ws[i].Cells[1, 2];
                tencere[i].ID = range.Value;
                range = ws[i].Cells[2, 2];
                tencere[i].Tip = range.Value;
                range = ws[i].Cells[3, 2];
                tencere[i].Uretici = range.Value;
                range = ws[i].Cells[4, 2];
                tencere[i].DimT = float.Parse((range.Value).ToString());
                range = ws[i].Cells[5, 2];
                tencere[i].DimN = float.Parse((range.Value).ToString());
                TencereList.Add(tencere[i]);
            }
            exl.DisplayAlerts = false;
            exl.UserControl = false;
            wb.Close(0);
            exl.Quit();
            
            updateListView();
        }

        private void ÇıkışToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Button3_Click_1(object sender, EventArgs e)
        {
            txt1 = textBox2.Text;
            txt2 = textBox3.Text;
            txt3 = textBox4.Text;
            txt4 = customTextBox1.Text;
            txt5 = customTextBox2.Text;
            NewTencere form1 = new NewTencere();
            form1.loadTencere = true;
            form1.Show();

        }

        private void Button8_Click(object sender, EventArgs e)
        {
            ReadExcel();
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void FormStartup_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }

        private void Button10_Click(object sender, EventArgs e)
        {

        }

        private void HakkındaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tabControl1.TabCount == 3) return;
            tabControl1.TabPages.Insert(2, tabPage3);
            tabControl1.SelectedTab = tabPage3;
            
        }

        private void Button10_Click_1(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabPage3);
            tabControl1.SelectedTab = tabPage1;
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(System.IO.Directory.GetParent(System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString() + @"\Kullanım Kılavuzu.docx");
        }


        private void Button7_Click(object sender, EventArgs e)
        {
            WriteExcel();

        }

        
        public static string secilenbobin;
        private void Button6_Click(object sender, EventArgs e)
        {
            string bobin;
            string defaulttype = "145";
            if (comboBox1.Text == String.Empty) 
            {
                bobin = defaulttype;
            }else
            {
                bobin = comboBox1.Text;
            }     
            secilenbobin = bobin;
            Graph graph = new Graph();
            if(graph.IsDisposed == true)
            {
                return;
            }else
            {
                graph.Show();
            }

            
        }

        private void YükleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Excelden tekrar yükleyecek.
            ReadExcel();
        }
        
    }
}
