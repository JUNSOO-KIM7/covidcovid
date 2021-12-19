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
using System.Net;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

using Excel = Microsoft.Office.Interop.Excel;



namespace covidcovid
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        private List<CountryData> CountryList = null;
        private List<CountryData> SelectedCountries = null;

        private Matrix Transform = null;
        private Matrix InverseTransform = null;
        private RectangleF WorldBounds;
        private PointF ClosePoint = new PointF(-1, -1);

        private CountryDataComparer Comparer =
           new CountryDataComparer(CountryDataComparer.CompareTypes.ByMaxCases);

        private bool IgnoreItemCheck = false;

        private void Form2_Load(object sender, EventArgs e)
        {
            LoadData();

            clbCountries.DataSource = CountryList;
            clbCountries.CheckOnClick = true;
        }
        private void LoadData()
        {
            string filename = "data" + DateTime.Now.ToString("yyyy_MM_dd") + ".csv";
            DownloadFile(filename);
            object[,] fields = LoadCsv(filename);
            CreateCountryData(fields);
        }
        private void CreateCountryData(object[,] fields)
        {
            Dictionary<string, CountryData> country_dict =
            new Dictionary<string, CountryData>();
            const int first_date_col = 5;
            int max_row = fields.GetUpperBound(0);
            int max_col = fields.GetUpperBound(1);
            int num_dates = max_col - first_date_col + 1;
            CountryData.Dates = new DateTime[num_dates];
            for (int col = 1; col <= num_dates; col++)
            {
                double double_value = (double)fields[1, col + first_date_col - 1];
                CountryData.Dates[col - 1] =
                DateTime.FromOADate(double_value);
            }

            const int country_col = 2;
            for (int country_num = 2; country_num <= max_row; country_num++)
            {
                string country_name = fields[country_num, country_col].ToString();

                CountryData country_data;

                if (country_dict.ContainsKey(country_name))
                {
                    country_data = country_dict[country_name];
                }
                else
                {
                    country_data = new CountryData();
                    country_data.Name = country_name;
                    country_data.Cases = new int[num_dates];
                    country_dict.Add(country_name, country_data);
                }

                for (int col = 1; col <= num_dates; col++)
                {
                    country_data.Cases[col - 1] +=
                        (int)(double)fields[country_num, col + first_date_col - 1];
                }
            }
            CountryList = country_dict.Values.ToList();

            foreach (CountryData country in CountryList)
            {
                country.SetMax();
            }

            CountryList.Sort(Comparer);
            for (int i = 0; i < CountryList.Count; i++)
            {
                CountryList[i].CountryNumber = i;
            }
        }

        private void DownloadFile(string filename)
        {
            if (!File.Exists(filename))
            {
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                try
                {
                    WebClient web_client = new WebClient();

                    const string url = "https://data.humdata.org/hxlproxy/api/data-preview.csv?url=https%3A%2F%2Fraw.githubusercontent.com%2FCSSEGISandData%2FCOVID-19%2Fmaster%2Fcsse_covid_19_data%2Fcsse_covid_19_time_series%2Ftime_series_covid19_confirmed_global.csv&filename=time_series_covid19_confirmed_global.csv";
                    web_client.DownloadFile(url, filename);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Download Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }
        private object[,] LoadCsv(string filename)
        {
            Excel.Application excel_app = new Excel.Application();

            filename = Application.StartupPath + "\\" + filename;
            Excel.Workbook workbook = excel_app.Workbooks.Open(
                filename,
                Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

            Excel.Range used_range = sheet.UsedRange;

            object[,] values = (object[,])used_range.Value2;

            workbook.Close(false, Type.Missing, Type.Missing);

            excel_app.Quit();

            return values;
        }
        private void clbCountries_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (IgnoreItemCheck) return;
            CountryData checked_country =
                clbCountries.Items[e.Index] as CountryData;
            SelectedCountries = GetCountryList(checked_country, e.NewValue);

            GraphCountries();
            SetTooltip(ClosePoint);
        }

        private List<CountryData> GetCountryList(
            CountryData checked_country, CheckState checked_state)
        {
            List<CountryData> country_list;
            if (clbCountries.CheckedItems.Count == 0)
            {
                country_list = new List<CountryData>();
            }
            else
            {
                country_list =
                    clbCountries.CheckedItems.Cast<CountryData>().ToList();
            }

            if (checked_country != null)
            {
                if (checked_state == CheckState.Checked)
                {
                    country_list.Add(checked_country);
                }
                else
                {
                    country_list.Remove(checked_country);
                }
            }
            return country_list;
        }

        private void GraphCountries()
        {
            ClosePoint = new PointF(-1, -1);
            if (SelectedCountries.Count == 0)
            {
                picGraph.Image = null;
                return;
            }

            float y_max = SelectedCountries.Max(country => country.Cases.Max());
            if (y_max < 10) y_max = 10;

            DefineTransform(SelectedCountries, y_max);

            Bitmap bm = new Bitmap(
                picGraph.ClientSize.Width,
                picGraph.ClientSize.Height);
            using (Graphics gr = Graphics.FromImage(bm))
            {
                gr.SmoothingMode = SmoothingMode.AntiAlias;
                gr.TextRenderingHint = TextRenderingHint.AntiAlias;
                gr.Transform = Transform;

                DrawAxes(gr);

                Color[] colors =
                {
                    Color.Red, Color.Green, Color.Blue, Color.Black,
                    Color.Cyan, Color.Orange,
                };
                int num_colors = colors.Length;
                using (Pen pen = new Pen(Color.Black, 0))
                {
                    foreach (CountryData country in SelectedCountries)
                    {
                        pen.Color = colors[country.CountryNumber % num_colors];
                        country.Draw(gr, pen, Transform);
                    }
                }
            }
            picGraph.Image = bm;
        }
        private void DefineTransform(List<CountryData> country_list, float y_max)
        {
            int num_cases = country_list[0].Cases.Length;
            WorldBounds = new RectangleF(0, 0, num_cases, y_max);
            int wid = picGraph.ClientSize.Width;
            int hgt = picGraph.ClientSize.Height - 1;
            const int margin = 4;
            PointF[] dest_points =
            {
                new PointF(margin, hgt - margin),
                new PointF(wid - margin, hgt - margin),
                new PointF(margin, margin),
            };
            Transform = new Matrix(WorldBounds, dest_points);
            InverseTransform = Transform.Clone();
            InverseTransform.Invert();
        }
        private void DrawAxes(Graphics gr)
        {
            using (Pen pen = new Pen(Color.Red, 0))
            {
                float y_max = WorldBounds.Bottom;
                int power = (int)Math.Log10(y_max);
                if (Math.Pow(10, power) > y_max / 2) power--;
                int y_step = (int)Math.Pow(10, power);
                if (y_step < 1) y_step = 1;

                gr.DrawLine(pen, 0, 0, 0, y_max);

                pen.Color = Color.Silver;
                float num_cases = WorldBounds.Right;
                for (int y = y_step; y < y_max; y += y_step)
                {
                    gr.DrawLine(pen, 0, y, num_cases, y);
                }

                GraphicsState state = gr.Save();
                gr.ResetTransform();
                using (Font font = new Font("Arial", 12, FontStyle.Regular))
                {
                    for (int y = y_step; y < y_max; y += y_step)
                    {
                        Point[] p = { new Point(0, y) };
                        Transform.TransformPoints(p);

                        gr.DrawString(y.ToString("n0"), font, Brushes.Black, p[0]);
                    }
                }
                gr.Restore(state);
                pen.Color = Color.Red;
                gr.DrawLine(pen, 0, 0, num_cases, 0);

                const int tick_y_pixels = 5;
                PointF[] tick_points = { new PointF(0, tick_y_pixels) };
                InverseTransform.TransformVectors(tick_points);
                float tick_y = -tick_points[0].Y;

                for (int i = 0; i < num_cases; i++)
                {
                    gr.DrawLine(pen, i, -tick_y, i, tick_y);
                }
            }
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            IgnoreItemCheck = true;
            for (int i = 0; i < clbCountries.Items.Count; i++)
                clbCountries.SetItemChecked(i, true);
            IgnoreItemCheck = false;
            RedrawGraph();
        }

        private void btnNone_Click(object sender, EventArgs e)
        {
            IgnoreItemCheck = true;
            for (int i = 0; i < clbCountries.Items.Count; i++)
                clbCountries.SetItemChecked(i, false);
            IgnoreItemCheck = false;
            RedrawGraph();
        }
        private void RedrawGraph()
        {
            SelectedCountries = GetCountryList(null, CheckState.Indeterminate);

            GraphCountries();
            
            SetTooltip(ClosePoint);
        }

        private void picGraph_MouseMove(object sender, MouseEventArgs e)
        {
            SetTooltip(e.Location);
        }

        private void SetTooltip(PointF point)
        {
            if (picGraph.Image == null) return;
            if (SelectedCountries == null) return;

            string new_tip = "";
            int day_num;
            int num_cases;
            foreach (CountryData country in SelectedCountries)
            {
                if (country.PointIsAt(point, out day_num,
                    out num_cases, out ClosePoint))
                {
                    new_tip = country.Name + "\n" +
                        CountryData.Dates[day_num].ToShortDateString() + "\n" +
                        num_cases.ToString("n0") + " cases";
                    break;
                }
            }

            //if (tipGraph.GetToolTip(picGraph) != new_tip)
               // tipGraph.SetToolTip(picGraph, new_tip);
            picGraph.Refresh();
        }

        private void picGraph_Paint(object sender, PaintEventArgs e)
        {
            if (ClosePoint.X < 0) return;

            const int radius = 3;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            float x = ClosePoint.X - radius;
            float y = ClosePoint.Y - radius;
            e.Graphics.FillEllipse(Brushes.White, x, y, 2 * radius, 2 * radius);
            e.Graphics.DrawEllipse(Pens.Red, x, y, 2 * radius, 2 * radius);
        }

        private void radSortByName_Click(object sender, EventArgs e)
        {
            if (CountryList == null) return;
            Comparer = new CountryDataComparer(CountryDataComparer.CompareTypes.ByName);
            clbCountries.DataSource = null;
            CountryList.Sort(Comparer);
            clbCountries.DataSource = CountryList;
            RedrawGraph();
        }

        private void radSortByMaxCases_Click(object sender, EventArgs e)
        {
            if (CountryList == null) return;
            Comparer = new CountryDataComparer(CountryDataComparer.CompareTypes.ByMaxCases);
            clbCountries.DataSource = null;
            CountryList.Sort(Comparer);
            clbCountries.DataSource = CountryList;
            RedrawGraph();
        }
    }
}
