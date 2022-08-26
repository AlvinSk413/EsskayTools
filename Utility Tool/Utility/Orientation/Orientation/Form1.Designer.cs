
namespace Orientation
{
    partial class frmOrientation
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            this.chkTSMemberType = new System.Windows.Forms.CheckBox();
            this.grpuseroption = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbPlane = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbRotation = new System.Windows.Forms.ComboBox();
            this.cmbDepth = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.PicBoxOrientaion = new System.Windows.Forms.PictureBox();
            this.chkErrorOnly = new System.Windows.Forms.CheckBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.chkUIOption = new System.Windows.Forms.CheckBox();
            this.pbar1 = new System.Windows.Forms.ProgressBar();
            this.chkontop = new System.Windows.Forms.CheckBox();
            this.lblsbar2 = new System.Windows.Forms.Label();
            this.lblsbar1 = new System.Windows.Forms.Label();
            this.btnexport = new System.Windows.Forms.Button();
            this.IsPrimary = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.XSMark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Phase = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Profile = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PartName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AtDepth = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Rotation = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OnPlane = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Remark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Guid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnclear = new System.Windows.Forms.Button();
            this.dgOrientation = new System.Windows.Forms.DataGridView();
            this.btncheck = new System.Windows.Forms.Button();
            this.btnmempos = new System.Windows.Forms.Button();
            this.grpuseroption.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PicBoxOrientaion)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgOrientation)).BeginInit();
            this.SuspendLayout();
            // 
            // chkTSMemberType
            // 
            this.chkTSMemberType.Location = new System.Drawing.Point(523, 393);
            this.chkTSMemberType.Name = "chkTSMemberType";
            this.chkTSMemberType.Size = new System.Drawing.Size(108, 20);
            this.chkTSMemberType.TabIndex = 45;
            this.chkTSMemberType.Text = "TS. Option";
            this.chkTSMemberType.UseVisualStyleBackColor = true;
            this.chkTSMemberType.CheckedChanged += new System.EventHandler(this.chkTSMemberType_CheckedChanged);
            // 
            // grpuseroption
            // 
            this.grpuseroption.Controls.Add(this.label1);
            this.grpuseroption.Controls.Add(this.cmbPlane);
            this.grpuseroption.Controls.Add(this.label2);
            this.grpuseroption.Controls.Add(this.cmbRotation);
            this.grpuseroption.Controls.Add(this.cmbDepth);
            this.grpuseroption.Controls.Add(this.label3);
            this.grpuseroption.Location = new System.Drawing.Point(523, 213);
            this.grpuseroption.Name = "grpuseroption";
            this.grpuseroption.Size = new System.Drawing.Size(110, 174);
            this.grpuseroption.TabIndex = 44;
            this.grpuseroption.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 21;
            this.label1.Text = "On Plane";
            // 
            // cmbPlane
            // 
            this.cmbPlane.FormattingEnabled = true;
            this.cmbPlane.Items.AddRange(new object[] {
            "MIDDLE",
            "RIGHT",
            "LEFT"});
            this.cmbPlane.Location = new System.Drawing.Point(6, 40);
            this.cmbPlane.Name = "cmbPlane";
            this.cmbPlane.Size = new System.Drawing.Size(98, 21);
            this.cmbPlane.TabIndex = 24;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 22;
            this.label2.Text = "Rotation";
            // 
            // cmbRotation
            // 
            this.cmbRotation.FormattingEnabled = true;
            this.cmbRotation.Items.AddRange(new object[] {
            "FRONT",
            "TOP",
            "BACK",
            "BELOW"});
            this.cmbRotation.Location = new System.Drawing.Point(6, 90);
            this.cmbRotation.Name = "cmbRotation";
            this.cmbRotation.Size = new System.Drawing.Size(98, 21);
            this.cmbRotation.TabIndex = 25;
            // 
            // cmbDepth
            // 
            this.cmbDepth.FormattingEnabled = true;
            this.cmbDepth.Items.AddRange(new object[] {
            "MIDDLE",
            "FRONT",
            "BEHIND"});
            this.cmbDepth.Location = new System.Drawing.Point(6, 140);
            this.cmbDepth.Name = "cmbDepth";
            this.cmbDepth.Size = new System.Drawing.Size(98, 21);
            this.cmbDepth.TabIndex = 26;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 119);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(49, 13);
            this.label3.TabIndex = 23;
            this.label3.Text = "At Depth";
            // 
            // PicBoxOrientaion
            // 
            this.PicBoxOrientaion.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.PicBoxOrientaion.Image = global::Orientation.Properties.Resources.Orientation;
            this.PicBoxOrientaion.Location = new System.Drawing.Point(2, 2);
            this.PicBoxOrientaion.Name = "PicBoxOrientaion";
            this.PicBoxOrientaion.Size = new System.Drawing.Size(515, 524);
            this.PicBoxOrientaion.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PicBoxOrientaion.TabIndex = 43;
            this.PicBoxOrientaion.TabStop = false;
            // 
            // chkErrorOnly
            // 
            this.chkErrorOnly.AutoSize = true;
            this.chkErrorOnly.Location = new System.Drawing.Point(523, 415);
            this.chkErrorOnly.Name = "chkErrorOnly";
            this.chkErrorOnly.Size = new System.Drawing.Size(72, 17);
            this.chkErrorOnly.TabIndex = 42;
            this.chkErrorOnly.Text = "Error Only";
            this.chkErrorOnly.UseVisualStyleBackColor = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Orientation.Properties.Resources.esskay;
            this.pictureBox1.Location = new System.Drawing.Point(523, 28);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(108, 162);
            this.pictureBox1.TabIndex = 39;
            this.pictureBox1.TabStop = false;
            // 
            // chkUIOption
            // 
            this.chkUIOption.AutoSize = true;
            this.chkUIOption.Location = new System.Drawing.Point(523, 196);
            this.chkUIOption.Name = "chkUIOption";
            this.chkUIOption.Size = new System.Drawing.Size(74, 17);
            this.chkUIOption.TabIndex = 41;
            this.chkUIOption.Text = "My Option";
            this.chkUIOption.UseVisualStyleBackColor = true;
            this.chkUIOption.CheckedChanged += new System.EventHandler(this.chkUIOption_CheckedChanged);
            // 
            // pbar1
            // 
            this.pbar1.Location = new System.Drawing.Point(2, 517);
            this.pbar1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pbar1.Name = "pbar1";
            this.pbar1.Size = new System.Drawing.Size(507, 10);
            this.pbar1.Step = 1;
            this.pbar1.TabIndex = 40;
            this.pbar1.Visible = false;
            // 
            // chkontop
            // 
            this.chkontop.AutoSize = true;
            this.chkontop.Checked = true;
            this.chkontop.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkontop.Location = new System.Drawing.Point(523, 2);
            this.chkontop.Name = "chkontop";
            this.chkontop.Size = new System.Drawing.Size(59, 17);
            this.chkontop.TabIndex = 38;
            this.chkontop.Text = "OnTop";
            this.chkontop.UseVisualStyleBackColor = true;
            // 
            // lblsbar2
            // 
            this.lblsbar2.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblsbar2.Location = new System.Drawing.Point(474, 529);
            this.lblsbar2.Name = "lblsbar2";
            this.lblsbar2.Size = new System.Drawing.Size(159, 20);
            this.lblsbar2.TabIndex = 37;
            this.lblsbar2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblsbar1
            // 
            this.lblsbar1.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblsbar1.Location = new System.Drawing.Point(10, 529);
            this.lblsbar1.Name = "lblsbar1";
            this.lblsbar1.Size = new System.Drawing.Size(445, 20);
            this.lblsbar1.TabIndex = 36;
            // 
            // btnexport
            // 
            this.btnexport.Location = new System.Drawing.Point(523, 467);
            this.btnexport.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnexport.Name = "btnexport";
            this.btnexport.Size = new System.Drawing.Size(100, 30);
            this.btnexport.TabIndex = 34;
            this.btnexport.Text = "Export";
            this.btnexport.UseVisualStyleBackColor = true;
            this.btnexport.Click += new System.EventHandler(this.btnexport_Click);
            // 
            // IsPrimary
            // 
            this.IsPrimary.HeaderText = "IsPrimary";
            this.IsPrimary.MinimumWidth = 6;
            this.IsPrimary.Name = "IsPrimary";
            this.IsPrimary.ReadOnly = true;
            this.IsPrimary.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.IsPrimary.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.IsPrimary.Width = 80;
            // 
            // XSMark
            // 
            this.XSMark.HeaderText = "Mark";
            this.XSMark.MinimumWidth = 6;
            this.XSMark.Name = "XSMark";
            this.XSMark.Width = 80;
            // 
            // Phase
            // 
            this.Phase.HeaderText = "Phase";
            this.Phase.MinimumWidth = 6;
            this.Phase.Name = "Phase";
            this.Phase.Width = 80;
            // 
            // Profile
            // 
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Profile.DefaultCellStyle = dataGridViewCellStyle7;
            this.Profile.HeaderText = "Profile";
            this.Profile.MinimumWidth = 6;
            this.Profile.Name = "Profile";
            this.Profile.Width = 125;
            // 
            // PartName
            // 
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PartName.DefaultCellStyle = dataGridViewCellStyle8;
            this.PartName.HeaderText = "Name";
            this.PartName.MinimumWidth = 6;
            this.PartName.Name = "PartName";
            this.PartName.Width = 125;
            // 
            // AtDepth
            // 
            this.AtDepth.HeaderText = "Depth";
            this.AtDepth.MinimumWidth = 6;
            this.AtDepth.Name = "AtDepth";
            this.AtDepth.ReadOnly = true;
            this.AtDepth.Width = 125;
            // 
            // Rotation
            // 
            this.Rotation.HeaderText = "Rotation";
            this.Rotation.MinimumWidth = 6;
            this.Rotation.Name = "Rotation";
            this.Rotation.ReadOnly = true;
            this.Rotation.Width = 125;
            // 
            // OnPlane
            // 
            this.OnPlane.HeaderText = "OnPlane";
            this.OnPlane.MinimumWidth = 6;
            this.OnPlane.Name = "OnPlane";
            this.OnPlane.ReadOnly = true;
            this.OnPlane.Width = 125;
            // 
            // Remark
            // 
            this.Remark.HeaderText = "Remark";
            this.Remark.MinimumWidth = 6;
            this.Remark.Name = "Remark";
            this.Remark.ReadOnly = true;
            this.Remark.Width = 125;
            // 
            // Guid
            // 
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Verdana", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Guid.DefaultCellStyle = dataGridViewCellStyle9;
            this.Guid.HeaderText = "GuidNo";
            this.Guid.MinimumWidth = 6;
            this.Guid.Name = "Guid";
            this.Guid.Visible = false;
            this.Guid.Width = 125;
            // 
            // btnclear
            // 
            this.btnclear.Location = new System.Drawing.Point(523, 497);
            this.btnclear.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnclear.Name = "btnclear";
            this.btnclear.Size = new System.Drawing.Size(100, 30);
            this.btnclear.TabIndex = 35;
            this.btnclear.Text = "Clear";
            this.btnclear.UseVisualStyleBackColor = true;
            this.btnclear.Click += new System.EventHandler(this.btnclear_Click);
            // 
            // dgOrientation
            // 
            this.dgOrientation.AllowUserToAddRows = false;
            this.dgOrientation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgOrientation.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Guid,
            this.Remark,
            this.OnPlane,
            this.Rotation,
            this.AtDepth,
            this.PartName,
            this.Profile,
            this.Phase,
            this.XSMark,
            this.IsPrimary});
            this.dgOrientation.Location = new System.Drawing.Point(2, 2);
            this.dgOrientation.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgOrientation.Name = "dgOrientation";
            this.dgOrientation.RowHeadersWidth = 51;
            this.dgOrientation.RowTemplate.Height = 24;
            this.dgOrientation.Size = new System.Drawing.Size(507, 524);
            this.dgOrientation.TabIndex = 33;
            this.dgOrientation.RowHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgOrientation_RowHeaderMouseDoubleClick);
            // 
            // btncheck
            // 
            this.btncheck.Location = new System.Drawing.Point(523, 437);
            this.btncheck.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btncheck.Name = "btncheck";
            this.btncheck.Size = new System.Drawing.Size(100, 30);
            this.btncheck.TabIndex = 32;
            this.btncheck.Text = "Check";
            this.btncheck.UseVisualStyleBackColor = true;
            this.btncheck.Click += new System.EventHandler(this.btncheck_Click);
            // 
            // btnmempos
            // 
            this.btnmempos.Location = new System.Drawing.Point(599, 191);
            this.btnmempos.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnmempos.Name = "btnmempos";
            this.btnmempos.Size = new System.Drawing.Size(32, 26);
            this.btnmempos.TabIndex = 46;
            this.btnmempos.Text = ">>";
            this.btnmempos.UseVisualStyleBackColor = true;
            this.btnmempos.Click += new System.EventHandler(this.btnmempos_Click);
            // 
            // frmOrientation
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(642, 550);
            this.Controls.Add(this.btnmempos);
            this.Controls.Add(this.chkTSMemberType);
            this.Controls.Add(this.grpuseroption);
            this.Controls.Add(this.PicBoxOrientaion);
            this.Controls.Add(this.chkErrorOnly);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.chkUIOption);
            this.Controls.Add(this.pbar1);
            this.Controls.Add(this.chkontop);
            this.Controls.Add(this.lblsbar2);
            this.Controls.Add(this.lblsbar1);
            this.Controls.Add(this.btnexport);
            this.Controls.Add(this.btnclear);
            this.Controls.Add(this.dgOrientation);
            this.Controls.Add(this.btncheck);
            this.MaximizeBox = false;
            this.Name = "frmOrientation";
            this.Text = "Orientation Checking Tool";
            this.Load += new System.EventHandler(this.frmOrientation_Load);
            this.grpuseroption.ResumeLayout(false);
            this.grpuseroption.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PicBoxOrientaion)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgOrientation)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkTSMemberType;
        private System.Windows.Forms.GroupBox grpuseroption;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbPlane;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbRotation;
        private System.Windows.Forms.ComboBox cmbDepth;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox PicBoxOrientaion;
        private System.Windows.Forms.CheckBox chkErrorOnly;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox chkUIOption;
        private System.Windows.Forms.ProgressBar pbar1;
        private System.Windows.Forms.CheckBox chkontop;
        private System.Windows.Forms.Label lblsbar2;
        private System.Windows.Forms.Label lblsbar1;
        private System.Windows.Forms.Button btnexport;
        private System.Windows.Forms.DataGridViewCheckBoxColumn IsPrimary;
        private System.Windows.Forms.DataGridViewTextBoxColumn XSMark;
        private System.Windows.Forms.DataGridViewTextBoxColumn Phase;
        private System.Windows.Forms.DataGridViewTextBoxColumn Profile;
        private System.Windows.Forms.DataGridViewTextBoxColumn PartName;
        private System.Windows.Forms.DataGridViewTextBoxColumn AtDepth;
        private System.Windows.Forms.DataGridViewTextBoxColumn Rotation;
        private System.Windows.Forms.DataGridViewTextBoxColumn OnPlane;
        private System.Windows.Forms.DataGridViewTextBoxColumn Remark;
        private System.Windows.Forms.DataGridViewTextBoxColumn Guid;
        private System.Windows.Forms.Button btnclear;
        private System.Windows.Forms.DataGridView dgOrientation;
        private System.Windows.Forms.Button btncheck;
        private System.Windows.Forms.Button btnmempos;
    }
}

