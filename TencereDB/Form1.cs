using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TencereDB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }
        public FormStartup formStartup = new FormStartup();
        IList<Tencere> liste = new List<Tencere>();
        public float[,] getValues()
        {
            float[,] finalvalues = new float[90, 6];
            //Seçilen texti açma işlemleri yapılır//
            return finalvalues;
        }

        public int bobinTip;
        public string filePath, fileName;
        private void Button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                filePath = openFileDialog1.FileName;
                fileName = openFileDialog1.SafeFileName;
                label8.Text = fileName;
                label8.Visible = true;
                
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //check if file still exist
            if(File.Exists(filePath))
            {
                foreach(Tencere element in liste)
                {
                    if (textBox1.Text == element.ID )
                    {

                    }
                }
                
                bobinTip = Convert.ToInt32(comboBox1.Text);
                //The file exists so do nothing
                this.Close();
            } else
            {
                MessageBox.Show("Dosya bulunamadı, tekrar deneyin.");

            }

            //Eğer ekliyeceğin tencerenin ID si zaten listede varsa 
            // bobin tipine bak, eklenilen bobin tipi daha önceden eklenmiş
            // listedeki tenceredede varsa uyar ( Bu tencerenin bu bobin tipi ile özellikleri 
            //daha önce eklenmiştir.) Eğer bobin tipi yoksa o bobin tipi ile güncelle.
            //than hide

        }
    }
}
