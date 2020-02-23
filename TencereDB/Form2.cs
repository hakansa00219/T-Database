using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TencereDB
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            updatePhotos();
        }
        void updatePhotos()
        {
            for (int i = 0; i < 2; i++)
            {
                Image[] asd = new Image[4];
                asd[i] = Image.FromFile(@"D:\TencereDB\TencereDB\photos\430.110.X.X" + @"\" + i.ToString() + ".jpg");
                // asd.SetResolution(asd.Width / 100, asd.Height / 100);
                ResizeImage(asd[i], 300, 300);
                if (i == 0) pictureBox2.Image = asd[i];
                if (i == 1) pictureBox3.Image = asd[i];
            }
        }
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            pictureBox2.Dispose();
            // Close();
            this.Dispose();
        }
    }
}
