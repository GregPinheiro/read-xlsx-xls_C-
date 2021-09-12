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

namespace ReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataTableCollection tableCollection;

        private void btRead_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" } )
            {
                if (ofd.ShowDialog().Equals(DialogResult.OK))
                {
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });

                            tableCollection = result.Tables;

                            cbPlan.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                cbPlan.Items.Add(table.TableName);
                            }
                        }
                    }
                }
            }
        }

        private void cbPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[cbPlan.SelectedItem.ToString()];
            dataGridView.DataSource = dt;
        }
    }
}
