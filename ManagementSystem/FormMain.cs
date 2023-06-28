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
        const string columnName = "판매자상품코드";
        string filePath = string.Empty;
        string newPath = string.Empty;

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            //[최종] 저장된 고객사별 code들 dataGridView에 불러오기
            //text로 불러오기. >> QQQ. 고객사가 수백개가 되기도 하나..?

            //일단은 수동입력으로 DataGridView add
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

        /// <summary> Whole Data담은 엑셀 파일 주소 불러오기 </summary>
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

        /// <summary> 고객사별 파일 추출하여 저장할 폴더 지정하기 </summary>
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

        /// <summary> tbSelectFile에서 지정한 통합 파일을 판매자 상품코드 별로 파일 분할 </summary>
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

                // 열 인덱스 찾기
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
                    // 지정된 열이 존재하지 않는 경우
                    MessageBox.Show($"열 {colName}이 존재하지 않습니다. 파일을 재확인해주세요.");
                    return;
                }


                //colName의 이름을 가진 열의 데이터들을 분할하여 Dictionary 저장
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

                //분할된 파일 생성
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