using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Text;
using System.Data.Common;
using System;

namespace ManagementSystem
{
    public partial class FormMain : Form
    {
        //기준열의 이름
        const string columnName = "판매자상품코드";
        //판매자 상품 코드들을 저장하는 .txt파일 위치
        const string pathCode = @"C:\Users\Hannah\Desktop\Test.txt";

        string filePath = string.Empty;
        string newPath = string.Empty;

        public FormMain()
        {
            InitializeComponent();
        }

        #region 판매자 관리 코드 불러오기
        private void FormMain_Load(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();

            LoadDataFromTxtToDataGridView(pathCode);
            AddRowNumbers();
        }

        /// <summary> .txt파일에서 판매자상품코드 정보들을 불러와 DataGridView에 뿌린다. </summary>
        /// <param name="filePath"> .txt파일이 저장된 경로 </param>
        private void LoadDataFromTxtToDataGridView(string filePath)
        {
            // .txt 파일에서 데이터 읽어오기
            string[] lines = File.ReadAllLines(filePath);

            for (int i = 0; i < lines.Length; i++)
            {
                string[] values = lines[i].Split('\t'); // 탭 구분자 사용
                dataGridView.Rows.Add(values);
            }
        }
        #endregion

        #region Export 섹션 버튼 이벤트
        /// <summary> Whole Data담은 엑셀 파일 주소 불러오기 </summary>
        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = tbSelectFile.Text = openFileDialog.FileName;
                //filePath = openFileDialog.FileName;
            }
        }

        /// <summary> 고객사별 파일 추출하여 저장할 폴더 지정하기 </summary>
        private void btnLocation_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                newPath = tbLocation.Text = folderBrowserDialog.SelectedPath;
                //newPath = folderBrowserDialog.SelectedPath;
            }
        }

        /// <summary> tbSelectFile에서 지정한 통합 파일을 판매자 상품코드 별로 파일 분할 </summary>
        private void btnExport_Click(object sender, EventArgs e)
        {
            SplitExcelFile(filePath, columnName);
        }
        #endregion

        public void SplitExcelFile(string sourceFilePath, string colName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(sourceFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; //0 : 1부터라고는 매뉴얼에 되어있지만, 왜인지 0으로 읽힘
                string formattedDateTime = string.Empty;

                // 열 이름으로 열 인덱스 찾기
                int columnIndex = -1;
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    string columnName = worksheet.Cells[1, col].Value?.ToString();
                    if (columnName == colName)
                    {
                        columnIndex = col;
                        break;
                    }
                }

                //해당 열의 값들을 앞 문자열 DISTINCT하여 list(columnValues)에 저장
                var columnValues = GetColumnValues(worksheet, columnIndex);

                //distinct한 data들을 기준으로 행을 읽고 Copy & New File 생성
                foreach (var value in columnValues)
                {
                    //value 값을 가진 행 index들 List로 가져오기
                    var rows = GetRowsByColumnValue(worksheet, columnIndex, value);

                    if (rows.Count > 0)
                    {
                        var newPackage = new ExcelPackage();
                        var newWorksheet = newPackage.Workbook.Worksheets.Add("Sheet1");

                        //저장된 도매처인 경우
                        if (CompareCodeData(value))
                        {
                            //Main View 'Content'의 내용을 구분자(,)로 구별하여 첫 행을 채워준다.
                            string[] cellValues = GetContentByCodeValue(value).Split(',');
                            for (int i = 0; i < cellValues.Length; i++)
                            {
                                newWorksheet.Cells[1, i + 1].Value = cellValues[i];
                            }

                            // 기존 엑셀 파일의 첫 번째 행과 새 엑셀 파일의 첫 번째 행을 비교하여 자리를 찾는다.
                            for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                            {
                                for (int col = 1; col <= newWorksheet.Dimension.Columns; col++)
                                {
                                    string existingCellValue = worksheet.Cells[1, column].Value?.ToString();
                                    string newCellValue = newWorksheet.Cells[1, col].Value?.ToString();

                                    // 일치하는 데이터가 존재하면 새로운 sheet의 열 기준으로 기존 셀 데이터를 복사합니다.
                                    if (existingCellValue == newCellValue)
                                    {
                                        // 해당 value의 row들을 돌면서 cell 복사
                                        int count = 2;
                                        string dataToCopy = string.Empty;
                                        for (int num = 0; num < rows.Count ; num++) //엑셀은 1부터
                                        {
                                            dataToCopy = worksheet.Cells[rows[num], column].Value?.ToString();
                                            newWorksheet.Cells[count, col].Value = dataToCopy;
                                            count++;
                                        }
                                    }
                                }
                            }
                        }
                        else //저장된 도매처가 아닌 경우 행을 그대로 복사한다.
                        {
                            //첫 행(컬럼명) 복사
                            string dataToCopy = string.Empty;
                            for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                            {
                                dataToCopy = worksheet.Cells[1, column].Value?.ToString();
                                newWorksheet.Cells[1, column].Value = dataToCopy;
                            }

                            //해당 value값의 rows 복사
                            for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                            {
                                int count = 2;
                                foreach (var row in rows)
                                {
                                    // 해당 value의 row들을 돌면서 cell 복사
                                    dataToCopy = worksheet.Cells[row, column].Value?.ToString();
                                    newWorksheet.Cells[count, column].Value = dataToCopy;
                                    count++;
                                }
                            }
                        }

                        #region 파일 저장
                        DateTime now = DateTime.Now;
                        formattedDateTime = now.ToString("yyMMdd_HHmm");

                        string newFilePath = Path.Combine(Path.GetDirectoryName(sourceFilePath), $"{value}_{formattedDateTime}.xlsx");
                        newPackage.SaveAs(new FileInfo(newFilePath));
                        #endregion
                    }
                }
            }
        }


        /// <summary> columnName과 이름이 일치하는 열의 data들을 Distinct하여 List로 출력한다. </summary>
        /// <param name="worksheet"></param>
        /// <returns> 해당 열의 data를 list로 출력 </returns>
        private List<string> GetColumnValues(ExcelWorksheet worksheet, int columnIndex)
        {
            // 해당 열의 첫 번째 행과 마지막 행 인덱스 가져오기
            int startRow = 2;  // 첫 번째 데이터 행의 인덱스
            int endRow = worksheet.Dimension.Rows;

            // 해당 열의 data들의 범위를 지정
            var columnRange = worksheet.Cells[startRow, columnIndex, endRow, columnIndex];
            
            // 해당 범위에서 cell값들을 Distinct처리하여 list(columnValues)에 넣어준다.
            var columnValues = new List<string>();
            foreach (var cell in columnRange)
            {
                string value = SplitString(cell.Value?.ToString());
                if (!string.IsNullOrEmpty(value) && !columnValues.Contains(value))
                {
                    columnValues.Add(value);
                }
            }
            return columnValues;
        }


        /// <summary> 문자열을 받아 앞부분 문자(char)만 남겨두고 제거 </summary>
        /// <param name="inputStr"> 문자열만 남겨두고 자를 문자열(도매처코드) </param>
        /// <returns> 문자열만 남겨진 문자열 </returns>
        public string SplitString(string inputStr)
        {
            int index = 0;
            foreach (var tmp in inputStr)
            {
                if (!Char.IsLetter(tmp))
                {
                    break;
                }
                index++; //char가 아닌 문자열의 Index
            }

            return inputStr.Substring(0, index);
        }


        /// <summary> 통합양식에서 Distinct한 data들이 ManagementSystem안에 저장되어 있는 데이터인지 확인하는 함수 </summary>
        /// <param name="codeData"> 비교하고자 하는 통합양식의 Distinct data </param>
        /// <returns> 저장되어 있는 코드일 경우 true를 리턴한다. </returns>
        private bool CompareCodeData(string codeData)
        {
            bool exists = false;
            string cellValue = string.Empty;

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                cellValue = row.Cells[1].Value?.ToString();
                if (cellValue == codeData)
                {
                    exists = true;
                    break;
                }
            }
            return exists;
        }


        /// <summary> 판매자 코드명에 맞는 Content열의 내용 불러오기 </summary>
        /// <param name="value"> 판매자 코드명 </param>
        /// <returns> 해당 코드 행의 Content열의 내용 </returns>
        private string GetContentByCodeValue(string value)
        {
            string content = string.Empty;

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                string code = row.Cells["colCode"].Value?.ToString();

                if (code == value)
                {
                    content = row.Cells["colContent"].Value?.ToString();
                    break;
                }
            }

            return content;
        }


        /// <summary> value값을 가진 행의 index들을 List로 가져온다. </summary>
        /// <param name="worksheet"> 원본 excel sheet </param>
        /// <param name="columnIndex"> 해당 열의 index </param>
        /// <param name="value"> distinct된 해당 열의 data </param>
        /// <returns> value(distinct data)값에 해당되는 행 index들을 List로 리턴 </returns>
        private List<int> GetRowsByColumnValue(ExcelWorksheet worksheet, int columnIndex, string value)
        {
            // 해당 열의 첫 번째 행과 마지막 행 인덱스 가져오기
            int startRow = 2;  // 첫 번째 데이터 행의 인덱스
            int endRow = worksheet.Dimension.Rows;

            var rows = new List<int>();

            // 해당 열을 기준으로 행(row)을 돌며 주어진 distinct data와 일치하면 해당 행 List(row)에 추가
            var columnRange = worksheet.Cells[startRow, columnIndex, endRow, columnIndex];
            for (int row = columnRange.Start.Row; row <= columnRange.End.Row; row++)
            {
                var cell = worksheet.Cells[row, columnRange.Start.Column];
                string cellValue = SplitString(cell.Value?.ToString());
                if (cellValue == value)
                {
                    rows.Add(row); //row index들을 list에 add
                }
            }
            return rows;
        }

        #region 판매자 코드 관리 섹션
        private void btnDelete_Click(object sender, EventArgs e)
        {
            EmptyText();
        }

        /// <summary> UI상에서 도매처 코드등록 그룹에 있는 TEXT들을 초기화한다. </summary>
        private void EmptyText()
        {
            tbCode.Text = string.Empty;
            tbName.Text = string.Empty;
            rtbContent.Text = string.Empty;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //int rowCount = dataGridView.RowCount;
            int rowIndex = dataGridView.Rows.Add();

            AddRowNumbers();
            dataGridView.Rows[rowIndex].Cells[1].Value = tbCode.Text;
            dataGridView.Rows[rowIndex].Cells[2].Value = tbName.Text;
            dataGridView.Rows[rowIndex].Cells[3].Value = rtbContent.Text;
            //dataGridView.Rows[rowIndex].Selected = true;

            EmptyText();
        }

        private void btnRemoveCode_Click(object sender, EventArgs e)
        {
            int index = dataGridView.CurrentCell.RowIndex;
            if (index >= 0)
            {
                // 선택된 행 삭제
                dataGridView.Rows.RemoveAt(index);
                AddRowNumbers();
            }
            else
            {
                MessageBox.Show("선택된 셀이 없습니다. 삭제할 행을 클릭하여 주십시오.");
            }
        }

        private void AddRowNumbers()
        {
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                dataGridView.Rows[i].Cells[0].Value = (i + 1).ToString();
            }
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            ExportDataToTxt(pathCode);
        }

        /// <summary> DataGridView에 있는 데이터들 텍스트 파일에 저장 </summary>
        /// <param name="filePath"> .txt </param>
        private void ExportDataToTxt(string filePath)
        {
            StringBuilder sb = new StringBuilder();

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    if (row.Cells[i].Value != null)
                    {
                        sb.Append(row.Cells[i].Value.ToString());
                    }
                    if (i < dataGridView.Columns.Count - 1)
                    {
                        sb.Append("\t"); // 탭 구분자 사용
                    }
                }
                sb.AppendLine();
            }

            // .txt 파일에 저장
            File.WriteAllText(filePath, sb.ToString());
        }
        #endregion
    }
}