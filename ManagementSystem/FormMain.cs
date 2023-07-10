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
        //���ؿ��� �̸�
        const string columnName = "�Ǹ��ڻ�ǰ�ڵ�";
        //�Ǹ��� ��ǰ �ڵ���� �����ϴ� .txt���� ��ġ
        const string pathCode = @"C:\Users\Hannah\Desktop\Test.txt";

        string filePath = string.Empty;
        string newPath = string.Empty;

        public FormMain()
        {
            InitializeComponent();
        }

        #region �Ǹ��� ���� �ڵ� �ҷ�����
        private void FormMain_Load(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();

            LoadDataFromTxtToDataGridView(pathCode);
            AddRowNumbers();
        }

        /// <summary> .txt���Ͽ��� �Ǹ��ڻ�ǰ�ڵ� �������� �ҷ��� DataGridView�� �Ѹ���. </summary>
        /// <param name="filePath"> .txt������ ����� ��� </param>
        private void LoadDataFromTxtToDataGridView(string filePath)
        {
            // .txt ���Ͽ��� ������ �о����
            string[] lines = File.ReadAllLines(filePath);

            for (int i = 0; i < lines.Length; i++)
            {
                string[] values = lines[i].Split('\t'); // �� ������ ���
                dataGridView.Rows.Add(values);
            }
        }
        #endregion

        #region Export ���� ��ư �̺�Ʈ
        /// <summary> Whole Data���� ���� ���� �ּ� �ҷ����� </summary>
        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = tbSelectFile.Text = openFileDialog.FileName;
                //filePath = openFileDialog.FileName;
            }
        }

        /// <summary> ���纰 ���� �����Ͽ� ������ ���� �����ϱ� </summary>
        private void btnLocation_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                newPath = tbLocation.Text = folderBrowserDialog.SelectedPath;
                //newPath = folderBrowserDialog.SelectedPath;
            }
        }

        /// <summary> tbSelectFile���� ������ ���� ������ �Ǹ��� ��ǰ�ڵ� ���� ���� ���� </summary>
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; //0 : 1���Ͷ��� �Ŵ��� �Ǿ�������, ������ 0���� ����
                string formattedDateTime = string.Empty;

                // �� �̸����� �� �ε��� ã��
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

                //�ش� ���� ������ �� ���ڿ� DISTINCT�Ͽ� list(columnValues)�� ����
                var columnValues = GetColumnValues(worksheet, columnIndex);

                //distinct�� data���� �������� ���� �а� Copy & New File ����
                foreach (var value in columnValues)
                {
                    //value ���� ���� �� index�� List�� ��������
                    var rows = GetRowsByColumnValue(worksheet, columnIndex, value);

                    if (rows.Count > 0)
                    {
                        var newPackage = new ExcelPackage();
                        var newWorksheet = newPackage.Workbook.Worksheets.Add("Sheet1");

                        //����� ����ó�� ���
                        if (CompareCodeData(value))
                        {
                            //Main View 'Content'�� ������ ������(,)�� �����Ͽ� ù ���� ä���ش�.
                            string[] cellValues = GetContentByCodeValue(value).Split(',');
                            for (int i = 0; i < cellValues.Length; i++)
                            {
                                newWorksheet.Cells[1, i + 1].Value = cellValues[i];
                            }

                            // ���� ���� ������ ù ��° ��� �� ���� ������ ù ��° ���� ���Ͽ� �ڸ��� ã�´�.
                            for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                            {
                                for (int col = 1; col <= newWorksheet.Dimension.Columns; col++)
                                {
                                    string existingCellValue = worksheet.Cells[1, column].Value?.ToString();
                                    string newCellValue = newWorksheet.Cells[1, col].Value?.ToString();

                                    // ��ġ�ϴ� �����Ͱ� �����ϸ� ���ο� sheet�� �� �������� ���� �� �����͸� �����մϴ�.
                                    if (existingCellValue == newCellValue)
                                    {
                                        // �ش� value�� row���� ���鼭 cell ����
                                        int count = 2;
                                        string dataToCopy = string.Empty;
                                        for (int num = 0; num < rows.Count ; num++) //������ 1����
                                        {
                                            dataToCopy = worksheet.Cells[rows[num], column].Value?.ToString();
                                            newWorksheet.Cells[count, col].Value = dataToCopy;
                                            count++;
                                        }
                                    }
                                }
                            }
                        }
                        else //����� ����ó�� �ƴ� ��� ���� �״�� �����Ѵ�.
                        {
                            //ù ��(�÷���) ����
                            string dataToCopy = string.Empty;
                            for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                            {
                                dataToCopy = worksheet.Cells[1, column].Value?.ToString();
                                newWorksheet.Cells[1, column].Value = dataToCopy;
                            }

                            //�ش� value���� rows ����
                            for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                            {
                                int count = 2;
                                foreach (var row in rows)
                                {
                                    // �ش� value�� row���� ���鼭 cell ����
                                    dataToCopy = worksheet.Cells[row, column].Value?.ToString();
                                    newWorksheet.Cells[count, column].Value = dataToCopy;
                                    count++;
                                }
                            }
                        }

                        #region ���� ����
                        DateTime now = DateTime.Now;
                        formattedDateTime = now.ToString("yyMMdd_HHmm");

                        string newFilePath = Path.Combine(Path.GetDirectoryName(sourceFilePath), $"{value}_{formattedDateTime}.xlsx");
                        newPackage.SaveAs(new FileInfo(newFilePath));
                        #endregion
                    }
                }
            }
        }


        /// <summary> columnName�� �̸��� ��ġ�ϴ� ���� data���� Distinct�Ͽ� List�� ����Ѵ�. </summary>
        /// <param name="worksheet"></param>
        /// <returns> �ش� ���� data�� list�� ��� </returns>
        private List<string> GetColumnValues(ExcelWorksheet worksheet, int columnIndex)
        {
            // �ش� ���� ù ��° ��� ������ �� �ε��� ��������
            int startRow = 2;  // ù ��° ������ ���� �ε���
            int endRow = worksheet.Dimension.Rows;

            // �ش� ���� data���� ������ ����
            var columnRange = worksheet.Cells[startRow, columnIndex, endRow, columnIndex];
            
            // �ش� �������� cell������ Distinctó���Ͽ� list(columnValues)�� �־��ش�.
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


        /// <summary> ���ڿ��� �޾� �պκ� ����(char)�� ���ܵΰ� ���� </summary>
        /// <param name="inputStr"> ���ڿ��� ���ܵΰ� �ڸ� ���ڿ�(����ó�ڵ�) </param>
        /// <returns> ���ڿ��� ������ ���ڿ� </returns>
        public string SplitString(string inputStr)
        {
            int index = 0;
            foreach (var tmp in inputStr)
            {
                if (!Char.IsLetter(tmp))
                {
                    break;
                }
                index++; //char�� �ƴ� ���ڿ��� Index
            }

            return inputStr.Substring(0, index);
        }


        /// <summary> ���վ�Ŀ��� Distinct�� data���� ManagementSystem�ȿ� ����Ǿ� �ִ� ���������� Ȯ���ϴ� �Լ� </summary>
        /// <param name="codeData"> ���ϰ��� �ϴ� ���վ���� Distinct data </param>
        /// <returns> ����Ǿ� �ִ� �ڵ��� ��� true�� �����Ѵ�. </returns>
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


        /// <summary> �Ǹ��� �ڵ�� �´� Content���� ���� �ҷ����� </summary>
        /// <param name="value"> �Ǹ��� �ڵ�� </param>
        /// <returns> �ش� �ڵ� ���� Content���� ���� </returns>
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


        /// <summary> value���� ���� ���� index���� List�� �����´�. </summary>
        /// <param name="worksheet"> ���� excel sheet </param>
        /// <param name="columnIndex"> �ش� ���� index </param>
        /// <param name="value"> distinct�� �ش� ���� data </param>
        /// <returns> value(distinct data)���� �ش�Ǵ� �� index���� List�� ���� </returns>
        private List<int> GetRowsByColumnValue(ExcelWorksheet worksheet, int columnIndex, string value)
        {
            // �ش� ���� ù ��° ��� ������ �� �ε��� ��������
            int startRow = 2;  // ù ��° ������ ���� �ε���
            int endRow = worksheet.Dimension.Rows;

            var rows = new List<int>();

            // �ش� ���� �������� ��(row)�� ���� �־��� distinct data�� ��ġ�ϸ� �ش� �� List(row)�� �߰�
            var columnRange = worksheet.Cells[startRow, columnIndex, endRow, columnIndex];
            for (int row = columnRange.Start.Row; row <= columnRange.End.Row; row++)
            {
                var cell = worksheet.Cells[row, columnRange.Start.Column];
                string cellValue = SplitString(cell.Value?.ToString());
                if (cellValue == value)
                {
                    rows.Add(row); //row index���� list�� add
                }
            }
            return rows;
        }

        #region �Ǹ��� �ڵ� ���� ����
        private void btnDelete_Click(object sender, EventArgs e)
        {
            EmptyText();
        }


        /// <summary> UI�󿡼� ����ó �ڵ��� �׷쿡 �ִ� TEXT���� �ʱ�ȭ�Ѵ�. </summary>
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
                // ���õ� �� ����
                dataGridView.Rows.RemoveAt(index);
                AddRowNumbers();
            }
            else
            {
                MessageBox.Show("���õ� ���� �����ϴ�. ������ ���� Ŭ���Ͽ� �ֽʽÿ�.");
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


        /// <summary> DataGridView�� �ִ� �����͵� �ؽ�Ʈ ���Ͽ� ���� </summary>
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
                        sb.Append("\t"); // �� ������ ���
                    }
                }
                sb.AppendLine();
            }

            // .txt ���Ͽ� ����
            File.WriteAllText(filePath, sb.ToString());
        }


        /// <summary> RichTextBox�� ���� �� �����͸� �ٿ��ֱ��ϴ� �Լ� </summary>
        private void PasteExcelDataToRichTextBox()
        {
            string clipboardText = Clipboard.GetText();

            if (!string.IsNullOrEmpty(clipboardText))
            {
                string[] lines = clipboardText.Split('\n');
                StringBuilder sb = new StringBuilder();

                foreach (string line in lines)
                {
                    string[] cells = line.Split('\t');

                    for (int i = 0; i < cells.Length; i++)
                    {
                        sb.Append(cells[i]);

                        if (i < cells.Length - 1)
                        {
                            sb.Append(",");
                        }
                    }

                    sb.AppendLine();
                }

                rtbContent.Text = sb.ToString();
            }
        }
        #endregion
    }
}