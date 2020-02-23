using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace TencereDB
{
    public partial class Graph : Form
    {
        public Graph()
        {
            InitializeComponent();
            this.Text = FormStartup.secilenbobin;
            
            

            
            ChartSpecs();
        }
        public listSelectedIndex a = new listSelectedIndex();
        public IList<Tencere> tlist = new List<Tencere>();
        public int bobintip = Convert.ToInt32(FormStartup.secilenbobin);
        public void ChartSpecs()
             
        {
            int index = a.returnIndex();
            tlist = FormStartup.TencereList;
            if (index == -1)
            {
                this.Close();
                //FormStartup formStartup = new FormStartup();
            } else
            {
                //CHART1 --------------------

                this.chart1.Series.Clear();
                int[] frequency = new int[91];
                for (int i = 10; i < 101; i++)
                {
                    frequency[i - 10] = i;
                }
                chart1.Series.Add("R");
                chart1.Series["R"].ChartType = SeriesChartType.Line;
                chart1.Series["R"].Color = Color.Red;
                chart1.Series["R"].BorderWidth = 3; chart1.Series["R"].Color = Color.Red;
                chart1.Series["R"].BorderWidth = 3;
                chart1.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
                try
                {
                    for (int k = 0; k < 91; k++)
                    {

                        switch (bobintip)
                        {
                            case 145:
                                chart1.Series["R"].Points.AddXY(frequency[k], tlist[index].specsClass[0][k].R);
                                break;
                            case 180:
                                chart1.Series["R"].Points.AddXY(frequency[k], tlist[index].specsClass[1][k].R);
                                break;
                            case 210:
                                chart1.Series["R"].Points.AddXY(frequency[k], tlist[index].specsClass[2][k].R);
                                break;
                            case 240:
                                chart1.Series["R"].Points.AddXY(frequency[k], tlist[index].specsClass[3][k].R);
                                break;
                            case 270:
                                chart1.Series["R"].Points.AddXY(frequency[k], tlist[index].specsClass[4][k].R);
                                break;
                        }

                    }
                }
                catch (Exception)
                {
  
                }
                

                // Chart2 -------------------------------
                this.chart2.Series.Clear();
                int[] frequency2 = new int[91];
               
                for (int i = 10; i < 101; i++)
                {
                    
                    frequency2[i - 10] = i;
                }
                chart2.Series.Add("L");
                chart2.Series["L"].ChartType = SeriesChartType.Line;
                chart2.ChartAreas[0].AxisY.LabelStyle.Format = "{0:0.##E+00}";
                chart2.Series["L"].Color = Color.Red;
                chart2.Series["L"].BorderWidth = 3; chart2.Series["L"].Color = Color.Red;
                chart2.Series["L"].BorderWidth = 3;
                chart2.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;

                try
                {
                    for (int k = 0; k < 91; k++)
                    {
                        switch (bobintip)
                        {
                            case 145:
                                chart2.Series["L"].Points.AddXY(frequency[k], tlist[index].specsClass[0][k].L);
                                break;
                            case 180:
                                chart2.Series["L"].Points.AddXY(frequency[k], tlist[index].specsClass[1][k].L);
                                break;
                            case 210:
                                chart2.Series["L"].Points.AddXY(frequency[k], tlist[index].specsClass[2][k].L);
                                break;
                            case 240:
                                chart2.Series["L"].Points.AddXY(frequency[k], tlist[index].specsClass[3][k].L);
                                break;
                            case 270:
                                chart2.Series["L"].Points.AddXY(frequency[k], tlist[index].specsClass[4][k].L);
                                break;
                        }
                    }
                    chart2.ChartAreas[0].RecalculateAxesScale();
                    chart2.ChartAreas[0].AxisY.Minimum = chart2.Series[0].Points.FindMinByValue().YValues[0] * 0.999;
                }
                catch(Exception)
                {

                }
                
                // Chart3 -------------------------------
                this.chart3.Series.Clear();
                int[] frequency3 = new int[91];
                
                for (int i = 10; i < 101; i++)
                {
                   
                    frequency3[i - 10] = i;
                }
                chart3.Series.Add("Z");
                chart3.Series["Z"].ChartType = SeriesChartType.Line;
                chart3.Series["Z"].Color = Color.Red;
                chart3.Series["Z"].BorderWidth = 3; chart3.Series["Z"].Color = Color.Red;
                chart3.Series["Z"].BorderWidth = 3;
                chart3.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
                try
                {
                    for (int k = 0; k < 91; k++)
                    {
                        switch (bobintip)
                        {
                            case 145:
                                chart3.Series["Z"].Points.AddXY(frequency[k], tlist[index].specsClass[0][k].Z);
                                break;
                            case 180:
                                chart3.Series["Z"].Points.AddXY(frequency[k], tlist[index].specsClass[1][k].Z);
                                break;
                            case 210:
                                chart3.Series["Z"].Points.AddXY(frequency[k], tlist[index].specsClass[2][k].Z);
                                break;
                            case 240:
                                chart3.Series["Z"].Points.AddXY(frequency[k], tlist[index].specsClass[3][k].Z);
                                break;
                            case 270:
                                chart3.Series["Z"].Points.AddXY(frequency[k], tlist[index].specsClass[4][k].Z);
                                break;
                        }
                    }

                }
                catch(Exception)
                {

                }
                
                // Chart4 ---------------------------------
                this.chart4.Series.Clear();
                int[] frequency4 = new int[91];
                
                for (int i = 10; i < 101; i++)
                {
                    
                    frequency4[i - 10] = i;
                }
                chart4.Series.Add("C");
                chart4.Series["C"].ChartType = SeriesChartType.Line;
                chart4.Series["C"].Color = Color.Red;
                chart4.Series["C"].BorderWidth = 3; chart4.Series["C"].Color = Color.Red;
                chart4.Series["C"].BorderWidth = 3;
                chart4.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
                try
                {
                    for (int k = 0; k < 91; k++)
                    {
                        switch (bobintip)
                        {
                            case 145:
                                chart4.Series["C"].Points.AddXY(frequency[k], tlist[index].specsClass[0][k].C);
                                break;
                            case 180:
                                chart4.Series["C"].Points.AddXY(frequency[k], tlist[index].specsClass[1][k].C);
                                break;
                            case 210:
                                chart4.Series["C"].Points.AddXY(frequency[k], tlist[index].specsClass[2][k].C);
                                break;
                            case 240:
                                chart4.Series["C"].Points.AddXY(frequency[k], tlist[index].specsClass[3][k].C);
                                break;
                            case 270:
                                chart4.Series["C"].Points.AddXY(frequency[k], tlist[index].specsClass[4][k].C);
                                break;
                        }
                    }
                }
                catch(Exception)
                {
                    MessageBox.Show("Seçilen tencerenin o bobin tipinde değerleri eklenmemiş.");
                    this.Close();
                }
                

                
            }

        }


        Point? prevPosition = null;
        ToolTip tooltip = new ToolTip();
        private void Chart1_MouseMove(object sender, MouseEventArgs e)
        {
            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            var results = chart1.HitTest(pos.X, pos.Y, false,
                                    ChartElementType.DataPoint);
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    var prop = result.Object as DataPoint;
                    if (prop != null)
                    {
                        var pointXPixel = result.ChartArea.AxisX.ValueToPixelPosition(prop.XValue);
                        var pointYPixel = result.ChartArea.AxisY.ValueToPixelPosition(prop.YValues[0]);

                        // check if the cursor is really close to the point (2 pixels around the point)
                        if (Math.Abs(pos.X - pointXPixel) < 4 &&
                            Math.Abs(pos.Y - pointYPixel) < 4)
                        {
                            tooltip.Show("X=" + prop.XValue + ", Y=" + prop.YValues[0], this.chart1,
                                            pos.X, pos.Y - 15);
                        }
                    }
                }
            }
        }

        private void Chart2_MouseMove(object sender, MouseEventArgs e)
        {
            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            var results = chart2.HitTest(pos.X, pos.Y, false,
                                    ChartElementType.DataPoint);
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    var prop = result.Object as DataPoint;
                    if (prop != null)
                    {
                        var pointXPixel = result.ChartArea.AxisX.ValueToPixelPosition(prop.XValue);
                        var pointYPixel = result.ChartArea.AxisY.ValueToPixelPosition(prop.YValues[0]);

                        // check if the cursor is really close to the point (2 pixels around the point)
                        if (Math.Abs(pos.X - pointXPixel) < 4 &&
                            Math.Abs(pos.Y - pointYPixel) < 4)
                        {
                            tooltip.Show("X=" + prop.XValue + ", Y=" + prop.YValues[0], this.chart1,
                                            pos.X, pos.Y - 15);
                        }
                    }
                }
            }
        }
        private void Chart3_MouseMove(object sender, MouseEventArgs e)
        {
            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            var results = chart3.HitTest(pos.X, pos.Y, false,
                                    ChartElementType.DataPoint);
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    var prop = result.Object as DataPoint;
                    if (prop != null)
                    {
                        var pointXPixel = result.ChartArea.AxisX.ValueToPixelPosition(prop.XValue);
                        var pointYPixel = result.ChartArea.AxisY.ValueToPixelPosition(prop.YValues[0]);

                        // check if the cursor is really close to the point (2 pixels around the point)
                        if (Math.Abs(pos.X - pointXPixel) < 4 &&
                            Math.Abs(pos.Y - pointYPixel) < 4)
                        {
                            tooltip.Show("X=" + prop.XValue + ", Y=" + prop.YValues[0], this.chart1,
                                            pos.X, pos.Y - 15);
                        }
                    }
                }
            }
        }
        private void Chart4_MouseMove(object sender, MouseEventArgs e)
        {
            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            var results = chart4.HitTest(pos.X, pos.Y, false,
                                    ChartElementType.DataPoint);
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    var prop = result.Object as DataPoint;
                    if (prop != null)
                    {
                        var pointXPixel = result.ChartArea.AxisX.ValueToPixelPosition(prop.XValue);
                        var pointYPixel = result.ChartArea.AxisY.ValueToPixelPosition(prop.YValues[0]);

                        // check if the cursor is really close to the point (2 pixels around the point)
                        if (Math.Abs(pos.X - pointXPixel) < 4 &&
                            Math.Abs(pos.Y - pointYPixel) < 4)
                        {
                            tooltip.Show("X=" + prop.XValue + ", Y=" + prop.YValues[0], this.chart1,
                                            pos.X, pos.Y - 15);
                        }
                    }
                }
            }
        }

    }
}
