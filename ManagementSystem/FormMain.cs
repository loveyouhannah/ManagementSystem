using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ManagementSystem
{
    public partial class FormMain : Form
    {
        string filePath = string.Empty;
        
        
        public FormMain()
        {
            InitializeComponent();
        }


        private void FormMain_Load(object sender, EventArgs e)
        {
            
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
                tbLocation.Text = folderBrowserDialog.SelectedPath;
            }
        }

        /// <summary> tbSelectFile���� ������ ���� ������ �Ǹ��� ��ǰ�ڵ� ���� ���� ���� </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {

        }


    }
}