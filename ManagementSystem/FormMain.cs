using System.Windows.Forms;
using System.IO;
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

            /*Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(tbSelectFile.Text);//@"C:\Users\Hannah\Desktop\testets.xlsx");
                worksheet = workbook.Worksheets.get_Item(1) as Worksheet;

                //방법 1 : Range 범위 지정
                //Microsoft.Office.Interop.Excel.Range range = worksheet.Range["B2", "K25"];
                //object[,] data = new object[25, 11];
                //range.Value = data;

                //방법 2 : Cell 개별 지정(1)
                //Range cell1 = worksheet.Cells[4, 4];
                //cell1.Value = "JOB NO";
                //Range cell2 = worksheet.Cells[5, 4];
                //cell2.Value = "LOT NO";

                //방법 3 : Cell 개별 지정(2)
                //((Range)worksheet.Cells[4, 4]).Value = "ZR1";
                //((Range)worksheet.Cells[5, 4]).Value = "ZR2";

                object[,] data1 = new object[,] { { "JOB NO" }, { "LOT NO" }, { "" }, { "UNIT NO" } };

                object[,] data2 = new object[,]{ { 0, 10, 12, 3, 4, 5},
                                                                { 1, 10, 12, 3, 4, 5},
                                                                { 2, 10, 12, 3, 4, 5},
                                                                { 3, 10, 12, 3, 4, 5},
                                                                { 4, 10, 12, 3, 4, 5},
                                                                { 5, 10, 12, 3, 4, 5},
                                                                { 6, 10, 12, 3, 4, 5},
                                                                { 7, 10, 12, 3, 4, 5} };

                Microsoft.Office.Interop.Excel.Range range = worksheet.Range["D4", "D7"];
                Microsoft.Office.Interop.Excel.Range range2 = worksheet.Range["E17", "J24"];

                range.Value = data1;
                range2.Value = data2;

                workbook.Save();

                workbook.Close();
                excelApp.Quit();

            }
            catch (Exception ex)
            {
                //  throw ex; //throw는 명시적으로 예외를 발생시킬 때 = 예외를 강제로 발 생시켜야하는 경우
                //throw하게되면 정상적인 프로그램 실행을 즉시 중단하고 가장 가까운 예외 처리기로 넘어간다.
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ReleaseExcelObject(worksheet);
                ReleaseExcelObject(workbook);
                ReleaseExcelObject(excelApp);

                MessageBox.Show("Saved!");
            }*/
        }


        public void SplitExcelFile(string sourceFilePath, string colName)
        {
            // 엑셀 애플리케이션 생성
            Excel.Application excelApp = new Excel.Application();

            try
            {
                // 원본 엑셀 파일 열기
                Workbook sourceWorkbook = excelApp.Workbooks.Open(sourceFilePath);
                Worksheet sourceWorksheet = sourceWorkbook.Worksheets.get_Item(1) as Worksheet;
                Excel.Range sourceRange = sourceWorksheet.UsedRange;

                // 특정 열의 데이터를 읽기 위한 열 인덱스 가져오기
                Excel.Range headerRow = (Excel.Range)sourceWorksheet.Rows[1];
                int columnIndex = GetColumnIndex(headerRow, colName);

                if (columnIndex == -1)
                {
                    // 주어진 열 이름을 찾을 수 없음
                    MessageBox.Show($"열 '{colName}'를 찾을 수 없습니다.");
                    return;
                }

                // 특정 열의 데이터 값으로 그룹화
                Dictionary<string, List<Excel.Range>> groups = new Dictionary<string, List<Excel.Range>>();

                for (int row = 2; row <= sourceRange.Rows.Count; row++)
                {
                    string value = ((Excel.Range)sourceRange.Cells[row, columnIndex]).Value?.ToString();

                    if (!string.IsNullOrEmpty(value))
                    {
                        if (!groups.ContainsKey(value))
                        {
                            groups[value] = new List<Excel.Range>();
                        }

                        //dictionary에 저장되어 있는 Key값(= value = 판매자 코드 명)에 따라 해당 행을 추가
                        groups[value].Add((Excel.Range)sourceRange.Rows[row]);
                    }
                }

                // 새로운 엑셀 파일로 데이터 복사
                string formattedDateTime = string.Empty;
                foreach (var group in groups)
                {
                    Excel.Workbook newWorkbook = excelApp.Workbooks.Add();
                    Excel.Worksheet newWorksheet = (Excel.Worksheet)newWorkbook.Worksheets[1];

                    int newRow = 1;
                    foreach (Excel.Range row in group.Value)
                    {
                        Excel.Range newRowRange = (Excel.Range)newWorksheet.Rows[newRow];
                        row.Copy(newRowRange);
                        newRow++;
                    }

                    // 새로운 엑셀 파일 저장
                    DateTime now = DateTime.Now;
                    formattedDateTime = now.ToString("yyMMdd_HHmm");

                    string newFilePath = Path.Combine(Path.GetDirectoryName(newPath), $"{group.Key}_{formattedDateTime}.xlsx");
                    newWorkbook.SaveAs(newFilePath);
                    newWorkbook.Close();
                }

                // 원본 엑셀 파일 닫기
                sourceWorkbook.Close();
            }
            catch (Exception ex)
            {
                // 예외 처리
                // 적절한 예외 처리를 수행하거나 오류 메시지를 표시하세요.
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // 엑셀 애플리케이션 종료
                excelApp.Quit();
            }
        }

        // 열 이름에 해당하는 열 인덱스 가져오기
        private int GetColumnIndex(Excel.Range headerRow, string columnName)
        {
            int columnIndex = -1;

            for (int column = 1; column <= headerRow.Columns.Count; column++)
            {
                string headerValue = ((Excel.Range)headerRow.Cells[1, column]).Value?.ToString();

                if (headerValue == columnName)
                {
                    columnIndex = column;
                    break;
                }
            }

            return columnIndex;
        }


        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    // 지정된 COM 개체와 연결된 지정된 RCW(RCW)의 참조 횟수를 감소시킵니다.
                    // obj : 해제할 COM 개체입니다.
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}