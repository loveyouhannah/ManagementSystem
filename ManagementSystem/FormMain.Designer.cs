namespace ManagementSystem
{
    partial class FormMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnExport = new Button();
            dataGridView = new DataGridView();
            colNo = new DataGridViewTextBoxColumn();
            colCode = new DataGridViewTextBoxColumn();
            colName = new DataGridViewTextBoxColumn();
            colContent = new DataGridViewTextBoxColumn();
            tbSelectFile = new TextBox();
            tbLocation = new TextBox();
            btnSelectFile = new Button();
            btnLocation = new Button();
            openFileDialog = new OpenFileDialog();
            folderBrowserDialog = new FolderBrowserDialog();
            gbExport = new GroupBox();
            gbManagement = new GroupBox();
            btnRemoveCode = new Button();
            gbEnroll = new GroupBox();
            btnAdd = new Button();
            btnDeleteText = new Button();
            lbContent = new Label();
            rtbContent = new RichTextBox();
            lbName = new Label();
            tbName = new TextBox();
            lbCode = new Label();
            tbCode = new TextBox();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            gbExport.SuspendLayout();
            gbManagement.SuspendLayout();
            gbEnroll.SuspendLayout();
            SuspendLayout();
            // 
            // btnExport
            // 
            btnExport.BackColor = Color.Brown;
            btnExport.FlatStyle = FlatStyle.Flat;
            btnExport.Font = new Font("Consolas", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            btnExport.ForeColor = SystemColors.ControlLight;
            btnExport.Location = new Point(649, 40);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(103, 88);
            btnExport.TabIndex = 0;
            btnExport.Text = "EXPORT";
            btnExport.UseVisualStyleBackColor = false;
            btnExport.Click += btnExport_Click;
            // 
            // dataGridView
            // 
            dataGridView.AllowUserToAddRows = false;
            dataGridView.BorderStyle = BorderStyle.None;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Columns.AddRange(new DataGridViewColumn[] { colNo, colCode, colName, colContent });
            dataGridView.Location = new Point(18, 48);
            dataGridView.Margin = new Padding(4, 5, 4, 5);
            dataGridView.Name = "dataGridView";
            dataGridView.RowHeadersVisible = false;
            dataGridView.RowHeadersWidth = 45;
            dataGridView.RowTemplate.Height = 25;
            dataGridView.Size = new Size(1561, 586);
            dataGridView.TabIndex = 1;
            // 
            // colNo
            // 
            colNo.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            colNo.FillWeight = 18.181818F;
            colNo.HeaderText = "No";
            colNo.MinimumWidth = 40;
            colNo.Name = "colNo";
            colNo.ReadOnly = true;
            colNo.Resizable = DataGridViewTriState.False;
            colNo.Width = 60;
            // 
            // colCode
            // 
            colCode.FillWeight = 140.909088F;
            colCode.HeaderText = "Code";
            colCode.MinimumWidth = 100;
            colCode.Name = "colCode";
            colCode.ReadOnly = true;
            colCode.Width = 150;
            // 
            // colName
            // 
            colName.FillWeight = 140.909088F;
            colName.HeaderText = "Name";
            colName.MinimumWidth = 100;
            colName.Name = "colName";
            colName.ReadOnly = true;
            colName.Width = 200;
            // 
            // colContent
            // 
            colContent.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            colContent.HeaderText = "Content";
            colContent.MinimumWidth = 500;
            colContent.Name = "colContent";
            colContent.ReadOnly = true;
            // 
            // tbSelectFile
            // 
            tbSelectFile.Location = new Point(27, 44);
            tbSelectFile.Margin = new Padding(4, 5, 4, 5);
            tbSelectFile.Name = "tbSelectFile";
            tbSelectFile.Size = new Size(478, 30);
            tbSelectFile.TabIndex = 2;
            // 
            // tbLocation
            // 
            tbLocation.Location = new Point(27, 94);
            tbLocation.Margin = new Padding(4, 5, 4, 5);
            tbLocation.Name = "tbLocation";
            tbLocation.Size = new Size(478, 30);
            tbLocation.TabIndex = 3;
            // 
            // btnSelectFile
            // 
            btnSelectFile.BackColor = Color.Black;
            btnSelectFile.FlatStyle = FlatStyle.Flat;
            btnSelectFile.ForeColor = Color.White;
            btnSelectFile.Location = new Point(521, 40);
            btnSelectFile.Name = "btnSelectFile";
            btnSelectFile.Size = new Size(116, 38);
            btnSelectFile.TabIndex = 4;
            btnSelectFile.Text = "변환파일";
            btnSelectFile.UseVisualStyleBackColor = false;
            btnSelectFile.Click += btnSelectFile_Click;
            // 
            // btnLocation
            // 
            btnLocation.BackColor = Color.Black;
            btnLocation.FlatStyle = FlatStyle.Flat;
            btnLocation.ForeColor = Color.White;
            btnLocation.Location = new Point(521, 91);
            btnLocation.Name = "btnLocation";
            btnLocation.Size = new Size(116, 38);
            btnLocation.TabIndex = 5;
            btnLocation.Text = "저장경로";
            btnLocation.UseVisualStyleBackColor = false;
            btnLocation.Click += btnLocation_Click;
            // 
            // openFileDialog
            // 
            openFileDialog.FileName = "openFileDialog";
            // 
            // gbExport
            // 
            gbExport.Controls.Add(tbSelectFile);
            gbExport.Controls.Add(btnLocation);
            gbExport.Controls.Add(btnExport);
            gbExport.Controls.Add(btnSelectFile);
            gbExport.Controls.Add(tbLocation);
            gbExport.Font = new Font("나눔고딕", 10F, FontStyle.Regular, GraphicsUnit.Point);
            gbExport.Location = new Point(838, 25);
            gbExport.Name = "gbExport";
            gbExport.Size = new Size(768, 147);
            gbExport.TabIndex = 6;
            gbExport.TabStop = false;
            gbExport.Text = "추출";
            // 
            // gbManagement
            // 
            gbManagement.Controls.Add(btnRemoveCode);
            gbManagement.Controls.Add(dataGridView);
            gbManagement.Font = new Font("나눔고딕", 10F, FontStyle.Regular, GraphicsUnit.Point);
            gbManagement.Location = new Point(12, 268);
            gbManagement.Name = "gbManagement";
            gbManagement.Padding = new Padding(14);
            gbManagement.Size = new Size(1593, 708);
            gbManagement.TabIndex = 7;
            gbManagement.TabStop = false;
            gbManagement.Text = "도매처 코드 관리";
            // 
            // btnRemoveCode
            // 
            btnRemoveCode.BackColor = Color.DimGray;
            btnRemoveCode.FlatStyle = FlatStyle.Flat;
            btnRemoveCode.Font = new Font("Consolas", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            btnRemoveCode.ForeColor = SystemColors.ControlLight;
            btnRemoveCode.Location = new Point(650, 645);
            btnRemoveCode.Name = "btnRemoveCode";
            btnRemoveCode.Size = new Size(254, 49);
            btnRemoveCode.TabIndex = 13;
            btnRemoveCode.Text = "REMOVE";
            btnRemoveCode.UseVisualStyleBackColor = false;
            btnRemoveCode.Click += btnRemoveCode_Click;
            // 
            // gbEnroll
            // 
            gbEnroll.Controls.Add(btnAdd);
            gbEnroll.Controls.Add(btnDeleteText);
            gbEnroll.Controls.Add(lbContent);
            gbEnroll.Controls.Add(rtbContent);
            gbEnroll.Controls.Add(lbName);
            gbEnroll.Controls.Add(tbName);
            gbEnroll.Controls.Add(lbCode);
            gbEnroll.Controls.Add(tbCode);
            gbEnroll.Font = new Font("나눔고딕", 10F, FontStyle.Regular, GraphicsUnit.Point);
            gbEnroll.Location = new Point(13, 25);
            gbEnroll.Name = "gbEnroll";
            gbEnroll.Padding = new Padding(10);
            gbEnroll.Size = new Size(805, 237);
            gbEnroll.TabIndex = 8;
            gbEnroll.TabStop = false;
            gbEnroll.Text = "도매처 코드 등록";
            // 
            // btnAdd
            // 
            btnAdd.BackColor = Color.Black;
            btnAdd.FlatStyle = FlatStyle.Flat;
            btnAdd.Font = new Font("Consolas", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            btnAdd.ForeColor = SystemColors.ControlLight;
            btnAdd.Location = new Point(674, 40);
            btnAdd.Name = "btnAdd";
            btnAdd.Size = new Size(108, 72);
            btnAdd.TabIndex = 12;
            btnAdd.Text = "ADD";
            btnAdd.UseVisualStyleBackColor = false;
            btnAdd.Click += btnAdd_Click;
            // 
            // btnDeleteText
            // 
            btnDeleteText.BackColor = Color.DimGray;
            btnDeleteText.FlatStyle = FlatStyle.Flat;
            btnDeleteText.Font = new Font("Consolas", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            btnDeleteText.ForeColor = SystemColors.ControlLight;
            btnDeleteText.Location = new Point(556, 40);
            btnDeleteText.Name = "btnDeleteText";
            btnDeleteText.Size = new Size(108, 72);
            btnDeleteText.TabIndex = 6;
            btnDeleteText.Text = "DELETE";
            btnDeleteText.UseVisualStyleBackColor = false;
            btnDeleteText.Click += btnDelete_Click;
            // 
            // lbContent
            // 
            lbContent.AutoSize = true;
            lbContent.Location = new Point(24, 130);
            lbContent.Name = "lbContent";
            lbContent.Size = new Size(96, 23);
            lbContent.TabIndex = 11;
            lbContent.Text = "Content :";
            // 
            // rtbContent
            // 
            rtbContent.BorderStyle = BorderStyle.FixedSingle;
            rtbContent.EnableAutoDragDrop = true;
            rtbContent.Location = new Point(127, 125);
            rtbContent.Name = "rtbContent";
            rtbContent.Size = new Size(655, 87);
            rtbContent.TabIndex = 10;
            rtbContent.Text = "";
            // 
            // lbName
            // 
            lbName.AutoSize = true;
            lbName.Location = new Point(43, 86);
            lbName.Name = "lbName";
            lbName.Size = new Size(77, 23);
            lbName.TabIndex = 9;
            lbName.Text = "Name :";
            // 
            // tbName
            // 
            tbName.Location = new Point(127, 82);
            tbName.Margin = new Padding(4, 5, 4, 5);
            tbName.Name = "tbName";
            tbName.Size = new Size(406, 30);
            tbName.TabIndex = 8;
            // 
            // lbCode
            // 
            lbCode.AutoSize = true;
            lbCode.Location = new Point(50, 45);
            lbCode.Name = "lbCode";
            lbCode.Size = new Size(70, 23);
            lbCode.TabIndex = 7;
            lbCode.Text = "Code :";
            // 
            // tbCode
            // 
            tbCode.Location = new Point(127, 41);
            tbCode.Margin = new Padding(4, 5, 4, 5);
            tbCode.Name = "tbCode";
            tbCode.Size = new Size(406, 30);
            tbCode.TabIndex = 6;
            // 
            // FormMain
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1621, 988);
            Controls.Add(gbEnroll);
            Controls.Add(gbManagement);
            Controls.Add(gbExport);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            Name = "FormMain";
            Text = "ManagementSystem";
            FormClosing += FormMain_FormClosing;
            Load += FormMain_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            gbExport.ResumeLayout(false);
            gbExport.PerformLayout();
            gbManagement.ResumeLayout(false);
            gbEnroll.ResumeLayout(false);
            gbEnroll.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Button btnExport;
        private DataGridView dataGridView;
        private TextBox tbSelectFile;
        private TextBox tbLocation;
        private Button btnSelectFile;
        private Button btnLocation;
        private OpenFileDialog openFileDialog;
        private FolderBrowserDialog folderBrowserDialog;
        private GroupBox gbExport;
        private GroupBox gbManagement;
        private GroupBox gbEnroll;
        private RichTextBox rtbContent;
        private Label lbName;
        private TextBox tbName;
        private Label lbCode;
        private TextBox tbCode;
        private DataGridViewTextBoxColumn colNo;
        private DataGridViewTextBoxColumn colCode;
        private DataGridViewTextBoxColumn colName;
        private DataGridViewTextBoxColumn colContent;
        private Label lbContent;
        private Button btnDeleteText;
        private Button btnAdd;
        private Button btnRemoveCode;
    }
}