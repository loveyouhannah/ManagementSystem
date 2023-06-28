using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Xml;

namespace ManagementSystem
{
    public partial class FormMain : Form
    {
        const string columnName = "�Ǹ��ڻ�ǰ�ڵ�";
        string filePath = string.Empty;
        string newPath = string.Empty;

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            //[����] ����� ���纰 code�� dataGridView�� �ҷ�����
            //text�� �ҷ�����. >> QQQ. ���簡 ���鰳�� �Ǳ⵵ �ϳ�..?

            //�ϴ��� �����Է����� DataGridView add
            dataGridView.Rows.Clear();

            //for (int i = 0; i < -.Count; i++)
            //{
            dataGridView.Rows.Add();
            dataGridView.Rows[0].Cells[colNo.Index].Value = 1;
            dataGridView.Rows[0].Cells[colName.Index].Value = "XR";
            dataGridView.Rows[0].Cells[colCode.Index].Value = "xrxr";

            dataGridView.Rows.Add();
            dataGridView.Rows[1].Cells[colNo.Index].Value = 2;
            dataGridView.Rows[1].Cells[colName.Index].Value = "ABC";
            dataGridView.Rows[1].Cells[colCode.Index].Value = "ABAB";

            dataGridView.Rows.Add();
            dataGridView.Rows[2].Cells[colNo.Index].Value = 3;
            dataGridView.Rows[2].Cells[colName.Index].Value = "GARDEN";
            dataGridView.Rows[2].Cells[colCode.Index].Value = "GD";
            //    dataGridView.Rows[i].Cells[dgvClassNo.Index].Value = i + 1;
            ////}
        }

        /// <summary> Whole Data���� ���� ���� �ּ� �ҷ����� </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = tbSelectFile.Text = openFileDialog.FileName;
                //filePath = openFileDialog.FileName;
            }
        }

        /// <summary> ���纰 ���� �����Ͽ� ������ ���� �����ϱ� </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLocation_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                newPath = tbLocation.Text = folderBrowserDialog.SelectedPath;
                //newPath = folderBrowserDialog.SelectedPath;
            }
        }

        /// <summary> tbSelectFile���� ������ ���� ������ �Ǹ��� ��ǰ�ڵ� ���� ���� ���� </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            SplitExcelFile(filePath, columnName);
        }

        public static void SplitExcelFile(string sourceFilePath, string colName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(sourceFilePath)))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                // �� �ε��� ã��
                int columnIndex = -1;
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    if (worksheet.Cells[1, col].Value?.ToString() == colName)
                    {
                        columnIndex = col;
                        break;
                    }
                }

                if (columnIndex == -1)
                {
                    // ������ ���� �������� �ʴ� ���
                    MessageBox.Show($"�� {colName}�� �������� �ʽ��ϴ�. ������ ��Ȯ�����ּ���.");
                    return;
                }


                //colName�� �̸��� ���� ���� �����͵��� �����Ͽ� Dictionary ����
                Dictionary<string, List<ExcelRange>> groups = new Dictionary<string, List<ExcelRange>>();
                for (int row = 2; row <= rowCount; row++)
                {
                    string value = worksheet.Cells[row, columnIndex].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        if (!groups.ContainsKey(value))
                        {
                            groups[value] = new List<ExcelRange>();
                        }
                        groups[value].Add(worksheet.Cells[row, columnIndex]);
                    }
                }

                //���ҵ� ���� ����
                string formattedDateTime = string.Empty;
                foreach (var kvp in groups)
                {
                    string groupName = kvp.Key;
                    List<ExcelRange> ranges = kvp.Value;

                    ExcelPackage newExcelPackage = new ExcelPackage();
                    ExcelWorksheet newWorksheet = newExcelPackage.Workbook.Worksheets.Add(groupName);

                    foreach (var range in ranges)
                    {
                        newWorksheet.Cells[range.Start.Row, range.Start.Column].Value = range.Value;
                    }

                    DateTime now = DateTime.Now;
                    formattedDateTime = now.ToString("yyMMdd_HHmm");

                    string newFilePath = Path.Combine(Path.GetDirectoryName(sourceFilePath), $"{groupName}_{formattedDateTime}.xlsx");
                    newExcelPackage.SaveAs(new FileInfo(newFilePath));
                }
            }
        }
    }
}