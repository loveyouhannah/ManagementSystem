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
            colName = new DataGridViewTextBoxColumn();
            colCode = new DataGridViewTextBoxColumn();
            colUse = new DataGridViewCheckBoxColumn();
            tbSelectFile = new TextBox();
            tbLocation = new TextBox();
            btnSelectFile = new Button();
            btnLocation = new Button();
            openFileDialog = new OpenFileDialog();
            folderBrowserDialog = new FolderBrowserDialog();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            SuspendLayout();
            // 
            // btnExport
            // 
            btnExport.BackColor = Color.DimGray;
            btnExport.FlatStyle = FlatStyle.Flat;
            btnExport.Font = new Font("Consolas", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            btnExport.ForeColor = SystemColors.ControlLight;
            btnExport.Location = new Point(632, 29);
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
            dataGridView.Columns.AddRange(new DataGridViewColumn[] { colNo, colName, colCode, colUse });
            dataGridView.Location = new Point(17, 140);
            dataGridView.Margin = new Padding(4, 5, 4, 5);
            dataGridView.Name = "dataGridView";
            dataGridView.RowHeadersVisible = false;
            dataGridView.RowHeadersWidth = 45;
            dataGridView.RowTemplate.Height = 25;
            dataGridView.Size = new Size(718, 612);
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
            colNo.Width = 80;
            // 
            // colName
            // 
            colName.FillWeight = 140.909088F;
            colName.HeaderText = "Name";
            colName.MinimumWidth = 8;
            colName.Name = "colName";
            colName.Width = 200;
            // 
            // colCode
            // 
            colCode.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            colCode.FillWeight = 140.909088F;
            colCode.HeaderText = "Code";
            colCode.MinimumWidth = 8;
            colCode.Name = "colCode";
            // 
            // colUse
            // 
            colUse.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            colUse.HeaderText = "Use";
            colUse.MinimumWidth = 40;
            colUse.Name = "colUse";
            colUse.Resizable = DataGridViewTriState.False;
            colUse.SortMode = DataGridViewColumnSortMode.Automatic;
            colUse.Width = 80;
            // 
            // tbSelectFile
            // 
            tbSelectFile.Location = new Point(17, 30);
            tbSelectFile.Margin = new Padding(4, 5, 4, 5);
            tbSelectFile.Name = "tbSelectFile";
            tbSelectFile.Size = new Size(478, 31);
            tbSelectFile.TabIndex = 2;
            // 
            // tbLocation
            // 
            tbLocation.Location = new Point(17, 80);
            tbLocation.Margin = new Padding(4, 5, 4, 5);
            tbLocation.Name = "tbLocation";
            tbLocation.Size = new Size(478, 31);
            tbLocation.TabIndex = 3;
            // 
            // btnSelectFile
            // 
            btnSelectFile.Location = new Point(511, 29);
            btnSelectFile.Name = "btnSelectFile";
            btnSelectFile.Size = new Size(116, 38);
            btnSelectFile.TabIndex = 4;
            btnSelectFile.Text = "변환파일";
            btnSelectFile.UseVisualStyleBackColor = true;
            btnSelectFile.Click += btnSelectFile_Click;
            // 
            // btnLocation
            // 
            btnLocation.Location = new Point(511, 79);
            btnLocation.Name = "btnLocation";
            btnLocation.Size = new Size(116, 38);
            btnLocation.TabIndex = 5;
            btnLocation.Text = "저장경로";
            btnLocation.UseVisualStyleBackColor = true;
            btnLocation.Click += btnLocation_Click;
            // 
            // openFileDialog
            // 
            openFileDialog.FileName = "openFileDialog";
            // 
            // FormMain
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(755, 772);
            Controls.Add(btnLocation);
            Controls.Add(btnSelectFile);
            Controls.Add(tbLocation);
            Controls.Add(tbSelectFile);
            Controls.Add(dataGridView);
            Controls.Add(btnExport);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            Name = "FormMain";
            Text = "ManagementSystem";
            Load += FormMain_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnExport;
        private DataGridView dataGridView;
        private DataGridViewTextBoxColumn colNo;
        private DataGridViewTextBoxColumn colName;
        private DataGridViewTextBoxColumn colCode;
        private DataGridViewCheckBoxColumn colUse;
        private TextBox tbSelectFile;
        private TextBox tbLocation;
        private Button btnSelectFile;
        private Button btnLocation;
        private OpenFileDialog openFileDialog;
        private FolderBrowserDialog folderBrowserDialog;
    }
}