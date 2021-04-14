using SfDataGrid_Demo;
using Syncfusion.WinForms.Controls;
using Syncfusion.WinForms.DataGrid;
using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;
using Syncfusion.WinForms.DataGrid.Enums;
using System.Drawing;
using Syncfusion.WinForms.DataGrid.Renderers;
using Syncfusion.WinForms.GridCommon.ScrollAxis;
using Syncfusion.Data.Extensions;
using Syncfusion.XlsIO;
using Syncfusion.WinForms.DataGridConverter;
using System.IO;

namespace SfDataGrid_Demo
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Constructor

        /// <summary>
        /// Initializes the new instance for the Form.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            sfDataGrid1.DataSource = new OrderInfoCollection().OrdersListDetails;
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {

            

            var options = new ExcelExportingOptions();
            options.CellExporting += OnCellExporting;

            options.ExcelVersion = ExcelVersion.Excel2013;
            var excelEngine = sfDataGrid1.ExportToExcel(sfDataGrid1.View, options);
            var workBook = excelEngine.Excel.Workbooks[0];

            SaveFileDialog saveFilterDialog = new SaveFileDialog
            {
                FilterIndex = 2,
                Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx|Excel 2013 File(*.xlsx)|*.xlsx"
            };

            if (saveFilterDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (Stream stream = saveFilterDialog.OpenFile())
                {
                    if (saveFilterDialog.FilterIndex == 1)
                        workBook.Version = ExcelVersion.Excel97to2003;
                    else if (saveFilterDialog.FilterIndex == 2)
                        workBook.Version = ExcelVersion.Excel2010;
                    else
                        workBook.Version = ExcelVersion.Excel2013;
                    workBook.SaveAs(stream);
                }

                //Message box confirmation to view the created workbook.
                if (MessageBox.Show(this.sfDataGrid1, "Do you want to view the workbook?", "Workbook has been created",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {

                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(saveFilterDialog.FileName);
                }
            }
        }

        private void OnCellExporting(object sender, Syncfusion.WinForms.DataGridConverter.Events.DataGridCellExcelExportingEventArgs e)
        {
            var record = e.NodeEntry as OrderInfo;

            // Style for OrderID column is changed based on the values in CustomerID column
            if(e.ColumnName == "OrderID")
            {
                if(record != null && record.CustomerID == "FRANS")
                {
                    e.Range.CellStyle.Color = Color.SkyBlue;
                    e.Range.CellStyle.Font.Color = ExcelKnownColors.Red;
                    e.Range.CellStyle.Font.FontName = "Arial";
                    e.Range.CellStyle.Font.Bold = true;
                }
            }
        }
    }
}
