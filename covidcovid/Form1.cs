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
using ExcelDataReader;

namespace covidcovid
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		DataTableCollection tablecollection;
		DataTable dt;

		private void BtnBrowse_Click(object sender, EventArgs e)
		{
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				CmbSheet.Items.Clear();
				CmbSheet.Text = "";
				txtFilename.Text = openFileDialog1.FileName;
				using (var stream = File.Open(openFileDialog1.FileName, FileMode.Open, FileAccess.Read))
				{
					if (Path.GetExtension(openFileDialog1.FileName).ToUpper() == ".XLS" || Path.GetExtension(openFileDialog1.FileName).ToUpper() == ".XLSX")
					{
						CmbSheet.Enabled = true;
						using (var reader = ExcelReaderFactory.CreateReader(stream))
						{
							var result = reader.AsDataSet(new ExcelDataSetConfiguration()
							{
								ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
								{
									EmptyColumnNamePrefix = "Column",
									UseHeaderRow = true
								}
							});
							dt = result.Tables[0];
							dataGridView1.DataSource = dt;
							tablecollection = result.Tables;
							foreach (DataTable table in tablecollection)
								CmbSheet.Items.Add(table.TableName); // add sheet to combobox
							CmbSheet.Enabled = true;
							CmbSheet.SelectedIndex = 0;
						}
					}
					else if (Path.GetExtension(openFileDialog1.FileName).ToUpper() == ".CSV" || Path.GetExtension(openFileDialog1.FileName).ToUpper() == ".TXT")
					{
						using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
						{
							var result = reader.AsDataSet(new ExcelDataSetConfiguration()
							{
								ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
								{
									EmptyColumnNamePrefix = "Column",
									UseHeaderRow = true
								}
							});
							CmbSheet.Enabled = false;
							dt = result.Tables[0];
							dataGridView1.DataSource = dt;
						}
					}

				}
			}
		}
		private void CmbSheet_SelectedIndexChanged(object sender, EventArgs e)
		{
			DataTable dt = tablecollection[CmbSheet.SelectedItem.ToString()];
			dataGridView1.DataSource = dt;
		}

		private void button1_Click(object sender, EventArgs e)
		{
			Form2 showForm2 = new Form2();
			showForm2.Show();
		}
	}
}
