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

            /*Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(tbSelectFile.Text);//@"C:\Users\Hannah\Desktop\testets.xlsx");
                worksheet = workbook.Worksheets.get_Item(1) as Worksheet;

                //��� 1 : Range ���� ����
                //Microsoft.Office.Interop.Excel.Range range = worksheet.Range["B2", "K25"];
                //object[,] data = new object[25, 11];
                //range.Value = data;

                //��� 2 : Cell ���� ����(1)
                //Range cell1 = worksheet.Cells[4, 4];
                //cell1.Value = "JOB NO";
                //Range cell2 = worksheet.Cells[5, 4];
                //cell2.Value = "LOT NO";

                //��� 3 : Cell ���� ����(2)
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
                //  throw ex; //throw�� ��������� ���ܸ� �߻���ų �� = ���ܸ� ������ �� �����Ѿ��ϴ� ���
                //throw�ϰԵǸ� �������� ���α׷� ������ ��� �ߴ��ϰ� ���� ����� ���� ó����� �Ѿ��.
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
            // ���� ���ø����̼� ����
            Excel.Application excelApp = new Excel.Application();

            try
            {
                // ���� ���� ���� ����
                Workbook sourceWorkbook = excelApp.Workbooks.Open(sourceFilePath);
                Worksheet sourceWorksheet = sourceWorkbook.Worksheets.get_Item(1) as Worksheet;
                Excel.Range sourceRange = sourceWorksheet.UsedRange;

                // Ư�� ���� �����͸� �б� ���� �� �ε��� ��������
                Excel.Range headerRow = (Excel.Range)sourceWorksheet.Rows[1];
                int columnIndex = GetColumnIndex(headerRow, colName);

                if (columnIndex == -1)
                {
                    // �־��� �� �̸��� ã�� �� ����
                    MessageBox.Show($"�� '{colName}'�� ã�� �� �����ϴ�.");
                    return;
                }

                // Ư�� ���� ������ ������ �׷�ȭ
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

                        //dictionary�� ����Ǿ� �ִ� Key��(= value = �Ǹ��� �ڵ� ��)�� ���� �ش� ���� �߰�
                        groups[value].Add((Excel.Range)sourceRange.Rows[row]);
                    }
                }

                // ���ο� ���� ���Ϸ� ������ ����
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

                    // ���ο� ���� ���� ����
                    DateTime now = DateTime.Now;
                    formattedDateTime = now.ToString("yyMMdd_HHmm");

                    string newFilePath = Path.Combine(Path.GetDirectoryName(newPath), $"{group.Key}_{formattedDateTime}.xlsx");
                    newWorkbook.SaveAs(newFilePath);
                    newWorkbook.Close();
                }

                // ���� ���� ���� �ݱ�
                sourceWorkbook.Close();
            }
            catch (Exception ex)
            {
                // ���� ó��
                // ������ ���� ó���� �����ϰų� ���� �޽����� ǥ���ϼ���.
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // ���� ���ø����̼� ����
                excelApp.Quit();
            }
        }

        // �� �̸��� �ش��ϴ� �� �ε��� ��������
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
                    // ������ COM ��ü�� ����� ������ RCW(RCW)�� ���� Ƚ���� ���ҽ�ŵ�ϴ�.
                    // obj : ������ COM ��ü�Դϴ�.
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