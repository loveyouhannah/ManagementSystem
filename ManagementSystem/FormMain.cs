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
                tbLocation.Text = folderBrowserDialog.SelectedPath;
            }
        }

        /// <summary> tbSelectFile에서 지정한 통합 파일을 판매자 상품코드 별로 파일 분할 </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {

        }


    }
}