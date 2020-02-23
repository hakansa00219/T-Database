using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace TencereDB
{
   
    public partial class NewTencere : Form
    {
        //dragging
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HTCAPTION = 0x2;
        [DllImport("User32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        //




        List<string> MessageQueue = new List<string>();
        public bool loadTencere = false;

        private SerialPort serialPort;
        string[] portNames;
        public IList<Tencere> _tencereList2 = new List<Tencere>();
        public NewTencere()
        {
            InitializeComponent();
            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorker1_RunWorkerCompleted;
        }
        public float[,] getValues()
        {
            float[,] finalvalues = new float[91, 4];
            //Seçilen texti açma işlemleri yapılır//
            return finalvalues;
        }


        public int bobinTip;
 

        private void Button1_Click(object sender, EventArgs e)
        {       
            if (_tencereList2.Count == 0)
            {
                MessageBox.Show("LCR'den gelen değerler boş." + "LCR meter kullanılmamış olabilir.");
            }
            else
            {
                _tencereList2[0].ID = textBox1.Text;
                _tencereList2[0].Tip = textBox2.Text;
                _tencereList2[0].Uretici = textBox3.Text;
                _tencereList2[0].DimT = float.Parse(customTextBox1.Text);
                _tencereList2[0].DimN = float.Parse(customTextBox2.Text);
            }
            if (comboBox1.Text == string.Empty)
            {
                MessageBox.Show("Bobin tipi seçilmemiş.");
            }else
            {
                bobinTip = Convert.ToInt32(comboBox1.Text);
            }
            bool exists = false;
           
            //check önceki listedeki tencereleri güncellemek (yeni bobin tipleri ile)için
            if (FormStartup.TencereList.Count != 0)
            {
                for (int a = 0; a<FormStartup.TencereList.Count; a++)
                {
                    try
                    {
                        if (textBox1.Text == FormStartup.TencereList[a].ID)
                        {
                            exists = true;
                            if (bobinTip == 145)
                            {
                                if (FormStartup.TencereList[a].specsClass[0][0] == null)
                                {
                                    //ise ekle

                                    for (int j = 0; j < 91; j++)
                                    {
                                        FormStartup.TencereList[a].specsClass[0][j] = (_tencereList2[0].specsClass[0][j]);
                                        FormStartup.TencereList[a].bobinTip = bobinTip;
                                    }


                                }
                                else
                                {
                                    //degilse 
                                    //MessageBox.Show("Daha önceden bu bobin tipinde ekleme yapılmış.");
                                }
                            }
                            else if (bobinTip == 180)
                            {
                                if (FormStartup.TencereList[a].specsClass[1][0] == null)
                                {
                                    //ise ekle

                                    for (int j = 0; j < 91; j++)
                                    {
                                        FormStartup.TencereList[a].specsClass[1][j] = (_tencereList2[0].specsClass[0][j]);
                                        FormStartup.TencereList[a].bobinTip = bobinTip;
                                    }

                                }
                                else
                                {
                                    //degilse 
                                    //MessageBox.Show("Daha önceden bu bobin tipinde ekleme yapılmış.");
                                }
                            }
                            else if (bobinTip == 210)
                            {
                                if (FormStartup.TencereList[a].specsClass[2][0] == null)
                                {
                                    //ise ekle

                                    for (int j = 0; j < 91; j++)
                                    {
                                        FormStartup.TencereList[a].specsClass[2][j] = (_tencereList2[0].specsClass[0][j]);
                                        FormStartup.TencereList[a].bobinTip = bobinTip;
                                    }

                                }
                                else
                                {
                                    //degilse 
                                    //MessageBox.Show("Daha önceden bu bobin tipinde ekleme yapılmış.");
                                }
                            }
                            else if (bobinTip == 240)
                            {
                                if (FormStartup.TencereList[a].specsClass[3][0] == null)
                                {
                                    //ise ekle

                                    for (int j = 0; j < 91; j++)
                                    {
                                        FormStartup.TencereList[a].specsClass[3][j] = (_tencereList2[0].specsClass[0][j]);
                                        FormStartup.TencereList[a].bobinTip = bobinTip;
                                    }

                                }
                                else
                                {
                                    //degilse 
                                    //MessageBox.Show("Daha önceden bu bobin tipinde ekleme yapılmış.");
                                }
                            }
                            else if (bobinTip == 270)
                            {
                                if (FormStartup.TencereList[a].specsClass[4][0] == null)
                                {
                                    //ise ekle

                                    for (int j = 0; j < 91; j++)
                                    {
                                        FormStartup.TencereList[a].specsClass[4][j] = (_tencereList2[0].specsClass[0][j]);
                                        FormStartup.TencereList[a].bobinTip = bobinTip;
                                    }

                                }
                                else
                                {
                                    //degilse 
                                    //MessageBox.Show("Daha önceden bu bobin tipinde ekleme yapılmış.");
                                }
                            }
                        }
                    } catch(NullReferenceException al)
                    {
                        // throw exeption
                    }
                }
            }
            if(exists == false)
            {
                FormStartup.TencereList.Add(_tencereList2[0]);
                FormStartup.TencereList[FormStartup.TencereList.Count - 1].bobinTip = bobinTip;
            }
            MessageBox.Show("eklendi");
            serialPort.Close();
            //The file exists so do nothing

            this.Close();

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            //serialPort = new SerialPort(portNames[comboBox2.SelectedIndex], 115200, Parity.None, 8, StopBits.One);
            openSerialPort();
        }


        private void openSerialPort()
        {
            if (serialPort == null)
            {
                // int a = 1 + comboBox4.SelectedIndex;
                serialPort = new SerialPort(portNames[comboBox2.SelectedIndex], 115200, Parity.None, 8, StopBits.One);
                //serialPort.ReadBufferSize = 9;
                serialPort.ReadTimeout = 50;
                serialPort.DiscardNull = true;
                serialPort.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
                serialPort.Open();
                serialPort.DiscardNull = false;
                label8.Text = "Port Opened";
                //serialPort.DtrEnable = dtr;

            } else if (serialPort.IsOpen )
                serialPort.Close();
        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            
            try
            {
                SerialPort sp = (SerialPort)sender;
                if (serialPort.BytesToRead != 0)
                {
                    byte[] buffer = new byte[200];
                    sp.Read(buffer, 0, serialPort.BytesToRead);

                    serialPort.DiscardInBuffer();
                    //Console.WriteLine( "onReading " + BitConverter.ToString(buffer));

                    Console.WriteLine("onReading " + ASCIIEncoding.ASCII.GetString(buffer).Trim(ArrTrim) );
                    String b = ASCIIEncoding.ASCII.GetString(buffer).Trim(ArrTrim);
                    if(b.Contains("\n"))
                        b = b.Substring(0, b.IndexOf("\n"));

                    String[] a = b.Split(',');
                    //a[3] = a[3].Replace('\n','');
                    //a[3] = a[3].Replace('\r', '');
                    decimal[] arr = new decimal[4];

                    for (int x = 0; x < 4; x++)
                    {
                        //a[x] = a[x].Replace(@"\r\n","");
                        arr[x] = decimal.Parse(a[x].Replace('.', ','), System.Globalization.NumberStyles.Number

                                                                      | System.Globalization.NumberStyles.AllowExponent);
                    }
                        

                    //Console.WriteLine("Z = " + arr[0].ToString() + "  Ls = " + arr[1].ToString() + " f = " + frequency.ToString());
                    _tencereList2[0].specsClass[selectedIndex][(frequency / 1000) - 10] = new TencereSpecs();
                    _tencereList2[0].specsClass[selectedIndex][(frequency / 1000) - 10].R = (double)arr[0];
                    _tencereList2[0].specsClass[selectedIndex][(frequency / 1000) - 10].L = (double)arr[1];
                    _tencereList2[0].specsClass[selectedIndex][(frequency / 1000) - 10].Z = (double)arr[2];
                    _tencereList2[0].specsClass[selectedIndex][(frequency / 1000) - 10].C = (double)arr[3];


                    if (sweepStarted)
                    {
                        
                        frequency += 1000;
                        if (frequency <= 100000)
                        {
                            
                            MessageQueue.Add(SystemCommands.setFrequency + frequency.ToString());
                            MessageQueue.Add(SystemCommands.measure);
                            progressBar1.Invoke((MethodInvoker)(() => progressBar1.Value = frequency / 1000)) ;

                        }
                        else
                        {
                            comboBox1.Invoke((MethodInvoker)(() => comboBox1.Enabled = true));
                            aTimer.Interval = 80;
                            MessageQueue.Add(SystemCommands.beepToneC);
                            MessageQueue.Add(SystemCommands.beep);
                            MessageQueue.Add(SystemCommands.beepToneB);
                            MessageQueue.Add(SystemCommands.beep);
                            //MessageBox.Show("Sweep bitti.");
                            sweepStarted = false;
                           // MessageQueue.Add(SystemCommands.closePort);


                            // this.Invoke(new MethodInvoker(this.Hide));

                        }

                    }

                    //AppConsoleWrite(false, buffer, null);
                }
            }
            catch (System.ArgumentOutOfRangeException a )
            {
                // do nothing
            } catch (System.IndexOutOfRangeException we)
            {

            }
        }
        char[] ArrTrim = {  (char)(0x00),
                            (char)(0x0A),
                            (char)(0x0D),
                            (char)(0x20),
        };

        private void NewTencere_Load(object sender, EventArgs e)
        {
            UpdateSerialPortMenu();
            button5.Enabled = true;
            textBox1.Text = FormStartup.txt1;
            textBox2.Text = FormStartup.txt2;
            textBox3.Text = FormStartup.txt3;
            customTextBox1.Text = FormStartup.txt4;
            customTextBox2.Text = FormStartup.txt5;
            // check porta bağlımı değilmi bağlıysa open yaz değilse closed yaz

            if (loadTencere)
            {
                Tencere ten = new Tencere();
                //.ıd = "1";
                // ten.tip = "f";
                ten = FormStartup.TencereList[FormStartup.selectedIndex];
                _tencereList2.Add(ten);
           
            } else
            {
                Tencere ten = new Tencere();
                //.ıd = "1";
               // ten.tip = "f";
                _tencereList2.Add(ten);
                _tencereList2[0].specsClass = new TencereSpecs[5][];
                for(int i = 0;i < 5; i++)
                {
                    _tencereList2[0].specsClass[i] = new TencereSpecs[91];
                }
            }



        }

        private void Button4_Click(object sender, EventArgs e)
        {
            UpdateSerialPortMenu();

        }
        private void UpdateSerialPortMenu()
        {
            comboBox2.Items.Clear();
            int a = 0;
            portNames = SerialPort.GetPortNames();     //<-- Reads all available comPorts
            foreach (var portName in portNames)
            {
                comboBox2.Items.Add(portName);                  //<-- Adds Ports to combobox
                a++;
            }
            if (a == 0)
                comboBox2.Items.Add("No Serial Port");
            comboBox2.SelectedIndex = 0;
        }
        
        private void Button5_Click(object sender, EventArgs e)
        {
            MessageQueue.Add(SystemCommands.init);
            MessageQueue.Add(SystemCommands.modeLCD);
            MessageQueue.Add(SystemCommands.averaging10);
            MessageQueue.Add(SystemCommands.rangeAuto);
            MessageQueue.Add(SystemCommands.setFrequency + frequency.ToString());
            MessageQueue.Add(SystemCommands.parameter1);
            MessageQueue.Add(SystemCommands.parameter2);
            MessageQueue.Add(SystemCommands.parameter3);
            MessageQueue.Add(SystemCommands.parameter4);
            MessageQueue.Add(SystemCommands.level);
            MessageQueue.Add(SystemCommands.constCurrent);
            MessageQueue.Add(SystemCommands.beepToneA);
            MessageQueue.Add(SystemCommands.beep);
            MessageQueue.Add(SystemCommands.beepToneB);
            MessageQueue.Add(SystemCommands.beep);
            SetTimer();
            //Tencere ten = new Tencere();
            //.ıd = "1";
            //ten.tip = "f";
            //_tencereList2.Add(ten);
            //_tencereList2[0].specsClass = new TencereSpecs[5][];
            //for(int i = 0;i < 5; i++)
            //{
            //    _tencereList2[0].specsClass[i] = new TencereSpecs[91];
            //}
            button5.Enabled = false;
            
        }

        bool sweepStarted = false;
        int frequency = 10000;
        int selectedIndex = 0;
        private void Button6_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == String.Empty)
            {
                MessageBox.Show("Önce bobin tipini seçin.");
                return;
            }
            sweepStarted = true;
            aTimer.Interval = 250;
            MessageQueue.Add(SystemCommands.beepToneB);
            MessageQueue.Add(SystemCommands.beep);
            MessageQueue.Add(SystemCommands.beepToneC);
            MessageQueue.Add(SystemCommands.beep);
            MessageQueue.Add(SystemCommands.measure);
            comboBox1.Enabled = false;
            // start the animation
            progressBar1.Visible = true;
            //progressBar1.Style = ProgressBarStyle.Marquee;
            // progressBar1.MarqueeAnimationSpeed = 1;

            selectedIndex = comboBox1.SelectedIndex;
            // start the job
            frequency = 10000;
            button6.Enabled = false;
            backgroundWorker1.RunWorkerAsync();
        }
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            waitThread();
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button6.Enabled = true;
            progressBar1.Visible = false;
        }

        private void waitThread()
        {
            while (sweepStarted);
        }

        private static System.Timers.Timer aTimer;
        private void SetTimer()
        {

            // Create a timer with a two second interval.
            aTimer = new System.Timers.Timer(80);
            aTimer.Elapsed += OnTimedEvent;
            aTimer.AutoReset = true;
            aTimer.Enabled = true;

        }
        private void sendPackageOverPort(string str)
        {
            if (serialPort != null)
            {

                if (serialPort.IsOpen)
                {
                    string a = str + SystemCommands. endlr;
                    byte[] arr = System.Text.Encoding.UTF8.GetBytes(a);
                    serialPort.Write(arr, 0, arr.Length); ;
                    // Console.WriteLine("sending " + a);
                }

            }

        }

        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            if (MessageQueue.Count != 0)
            {
                sendPackageOverPort(MessageQueue[0]);
                MessageQueue.RemoveAt(0);
            }
        }


        private void NewTencere_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort != null)
                if (serialPort.IsOpen) serialPort.Close();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            Close();
        }


        private void Panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }

        private void Panel2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }
    }
}
