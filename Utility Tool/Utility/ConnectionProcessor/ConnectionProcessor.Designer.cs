
namespace ConnectionProcessor
{
    partial class frmconnectionprocess
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmconnectionprocess));
            this.lblsbar1 = new System.Windows.Forms.Label();
            this.lblsbar2 = new System.Windows.Forms.Label();
            this.lblxsuser = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tbpgconn = new System.Windows.Forms.TabPage();
            this.tvwpreviewconn = new Tekla.Structures.Dialog.UIControls.Tree();
            this.dgvconnectioncompare = new System.Windows.Forms.DataGridView();
            this.tbpgconnupdate = new System.Windows.Forms.TabPage();
            this.btnsetallattibuteinfo = new System.Windows.Forms.Button();
            this.btnsetattributes = new System.Windows.Forms.Button();
            this.btngetattributes = new System.Windows.Forms.Button();
            this.btnhlp = new System.Windows.Forms.Button();
            this.btngetallattibuteinfo = new System.Windows.Forms.Button();
            this.dgattibutedata = new System.Windows.Forms.DataGridView();
            this.TSUpdate = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.AttributeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AttributeType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AttributeValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label2 = new System.Windows.Forms.Label();
            this.txtattfieldvalue = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtattfieldname = new System.Windows.Forms.TextBox();
            this.tbpgview = new System.Windows.Forms.TabPage();
            this.tvwpreviewconntype1 = new Tekla.Structures.Dialog.UIControls.Tree();
            this.pbar1 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.btntvwtoggle = new System.Windows.Forms.Button();
            this.btnselect = new System.Windows.Forms.Button();
            this.btnselectconnection = new System.Windows.Forms.Button();
            this.rtxtgradechange = new System.Windows.Forms.RichTextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.txtimpfeet = new System.Windows.Forms.TextBox();
            this.txtimpinch = new System.Windows.Forms.TextBox();
            this.txtmm = new System.Windows.Forms.TextBox();
            this.btngrouping = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnsave = new System.Windows.Forms.Button();
            this.btnload = new System.Windows.Forms.Button();
            this.chkexist = new System.Windows.Forms.CheckBox();
            this.chkonTop = new System.Windows.Forms.CheckBox();
            this.chkIgnoreConnectionCode = new System.Windows.Forms.CheckBox();
            this.chkselgrp = new System.Windows.Forms.CheckBox();
            this.lblguid = new System.Windows.Forms.Label();
            this.btngrpguid = new System.Windows.Forms.Button();
            this.btnclear = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tabControl1.SuspendLayout();
            this.tbpgconn.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvconnectioncompare)).BeginInit();
            this.tbpgconnupdate.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgattibutedata)).BeginInit();
            this.tbpgview.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // lblsbar1
            // 
            this.lblsbar1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblsbar1.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblsbar1.Location = new System.Drawing.Point(8, 400);
            this.lblsbar1.Name = "lblsbar1";
            this.lblsbar1.Size = new System.Drawing.Size(361, 13);
            this.lblsbar1.TabIndex = 43;
            this.lblsbar1.Text = "Remarks";
            this.lblsbar1.Click += new System.EventHandler(this.lblsbar1_Click);
            // 
            // lblsbar2
            // 
            this.lblsbar2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblsbar2.ForeColor = System.Drawing.Color.Red;
            this.lblsbar2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblsbar2.Location = new System.Drawing.Point(467, 402);
            this.lblsbar2.Name = "lblsbar2";
            this.lblsbar2.Size = new System.Drawing.Size(325, 13);
            this.lblsbar2.TabIndex = 44;
            this.lblsbar2.Text = "Remarks";
            // 
            // lblxsuser
            // 
            this.lblxsuser.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblxsuser.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblxsuser.ForeColor = System.Drawing.Color.Blue;
            this.lblxsuser.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.lblxsuser.Location = new System.Drawing.Point(851, 400);
            this.lblxsuser.Name = "lblxsuser";
            this.lblxsuser.Size = new System.Drawing.Size(77, 12);
            this.lblxsuser.TabIndex = 87;
            this.lblxsuser.Text = "User";
            this.lblxsuser.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tbpgconn);
            this.tabControl1.Controls.Add(this.tbpgconnupdate);
            this.tabControl1.Controls.Add(this.tbpgview);
            this.tabControl1.Location = new System.Drawing.Point(6, 1);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(835, 398);
            this.tabControl1.TabIndex = 94;
            // 
            // tbpgconn
            // 
            this.tbpgconn.Controls.Add(this.tvwpreviewconn);
            this.tbpgconn.Controls.Add(this.dgvconnectioncompare);
            this.tbpgconn.Location = new System.Drawing.Point(4, 21);
            this.tbpgconn.Name = "tbpgconn";
            this.tbpgconn.Padding = new System.Windows.Forms.Padding(3);
            this.tbpgconn.Size = new System.Drawing.Size(827, 373);
            this.tbpgconn.TabIndex = 2;
            this.tbpgconn.Text = "Grouping";
            this.tbpgconn.UseVisualStyleBackColor = true;
            // 
            // tvwpreviewconn
            // 
            this.tvwpreviewconn.BackColor = System.Drawing.Color.White;
            this.tvwpreviewconn.CheckBoxes = true;
            this.tvwpreviewconn.Location = new System.Drawing.Point(582, 5);
            this.tvwpreviewconn.Name = "tvwpreviewconn";
            this.tvwpreviewconn.PathSeparator = "|";
            this.tvwpreviewconn.Size = new System.Drawing.Size(239, 361);
            this.tvwpreviewconn.TabIndex = 103;
            this.tvwpreviewconn.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvwpreviewconn_AfterSelect);
            this.tvwpreviewconn.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvwpreviewconn_NodeMouseDoubleClick);
            // 
            // dgvconnectioncompare
            // 
            this.dgvconnectioncompare.AllowUserToAddRows = false;
            this.dgvconnectioncompare.AllowUserToOrderColumns = true;
            this.dgvconnectioncompare.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllHeaders;
            this.dgvconnectioncompare.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvconnectioncompare.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dgvconnectioncompare.Location = new System.Drawing.Point(6, 6);
            this.dgvconnectioncompare.Name = "dgvconnectioncompare";
            this.dgvconnectioncompare.RowHeadersWidth = 51;
            this.dgvconnectioncompare.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvconnectioncompare.Size = new System.Drawing.Size(570, 360);
            this.dgvconnectioncompare.TabIndex = 0;
            this.dgvconnectioncompare.Tag = "9";
            this.dgvconnectioncompare.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvconnectioncompare_CellContentDoubleClick);
            this.dgvconnectioncompare.ColumnHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvconnectioncompare_ColumnHeaderMouseDoubleClick);
            // 
            // tbpgconnupdate
            // 
            this.tbpgconnupdate.Controls.Add(this.btnsetallattibuteinfo);
            this.tbpgconnupdate.Controls.Add(this.btnsetattributes);
            this.tbpgconnupdate.Controls.Add(this.btngetattributes);
            this.tbpgconnupdate.Controls.Add(this.btnhlp);
            this.tbpgconnupdate.Controls.Add(this.btngetallattibuteinfo);
            this.tbpgconnupdate.Controls.Add(this.dgattibutedata);
            this.tbpgconnupdate.Controls.Add(this.label2);
            this.tbpgconnupdate.Controls.Add(this.txtattfieldvalue);
            this.tbpgconnupdate.Controls.Add(this.label1);
            this.tbpgconnupdate.Controls.Add(this.txtattfieldname);
            this.tbpgconnupdate.Location = new System.Drawing.Point(4, 22);
            this.tbpgconnupdate.Name = "tbpgconnupdate";
            this.tbpgconnupdate.Padding = new System.Windows.Forms.Padding(3);
            this.tbpgconnupdate.Size = new System.Drawing.Size(827, 372);
            this.tbpgconnupdate.TabIndex = 1;
            this.tbpgconnupdate.Text = "Get/Set Connection";
            this.tbpgconnupdate.UseVisualStyleBackColor = true;
            // 
            // btnsetallattibuteinfo
            // 
            this.btnsetallattibuteinfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnsetallattibuteinfo.Location = new System.Drawing.Point(719, 3);
            this.btnsetallattibuteinfo.Name = "btnsetallattibuteinfo";
            this.btnsetallattibuteinfo.Size = new System.Drawing.Size(80, 26);
            this.btnsetallattibuteinfo.TabIndex = 101;
            this.btnsetallattibuteinfo.Text = "Set All";
            this.btnsetallattibuteinfo.UseVisualStyleBackColor = true;
            this.btnsetallattibuteinfo.Click += new System.EventHandler(this.btnsetallattibuteinfo_Click);
            // 
            // btnsetattributes
            // 
            this.btnsetattributes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnsetattributes.Location = new System.Drawing.Point(560, 3);
            this.btnsetattributes.Name = "btnsetattributes";
            this.btnsetattributes.Size = new System.Drawing.Size(59, 26);
            this.btnsetattributes.TabIndex = 100;
            this.btnsetattributes.Text = "Set";
            this.btnsetattributes.UseVisualStyleBackColor = true;
            this.btnsetattributes.Click += new System.EventHandler(this.btnsetattributes_Click);
            // 
            // btngetattributes
            // 
            this.btngetattributes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btngetattributes.Location = new System.Drawing.Point(491, 3);
            this.btngetattributes.Name = "btngetattributes";
            this.btngetattributes.Size = new System.Drawing.Size(59, 26);
            this.btngetattributes.TabIndex = 99;
            this.btngetattributes.Text = "Get";
            this.btngetattributes.UseVisualStyleBackColor = true;
            this.btngetattributes.Click += new System.EventHandler(this.btngetattributes_Click);
            // 
            // btnhlp
            // 
            this.btnhlp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnhlp.Location = new System.Drawing.Point(809, 3);
            this.btnhlp.Name = "btnhlp";
            this.btnhlp.Size = new System.Drawing.Size(18, 26);
            this.btnhlp.TabIndex = 98;
            this.btnhlp.Text = "?";
            this.btnhlp.UseVisualStyleBackColor = true;
            // 
            // btngetallattibuteinfo
            // 
            this.btngetallattibuteinfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btngetallattibuteinfo.Location = new System.Drawing.Point(629, 3);
            this.btngetallattibuteinfo.Name = "btngetallattibuteinfo";
            this.btngetallattibuteinfo.Size = new System.Drawing.Size(80, 26);
            this.btngetallattibuteinfo.TabIndex = 97;
            this.btngetallattibuteinfo.Text = "Get All";
            this.btngetallattibuteinfo.UseVisualStyleBackColor = true;
            this.btngetallattibuteinfo.Click += new System.EventHandler(this.btngetallattibuteinfo_Click);
            // 
            // dgattibutedata
            // 
            this.dgattibutedata.AllowUserToAddRows = false;
            this.dgattibutedata.AllowUserToDeleteRows = false;
            this.dgattibutedata.AllowUserToResizeColumns = false;
            this.dgattibutedata.AllowUserToResizeRows = false;
            this.dgattibutedata.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgattibutedata.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.TSUpdate,
            this.AttributeName,
            this.AttributeType,
            this.AttributeValue});
            this.dgattibutedata.Location = new System.Drawing.Point(13, 32);
            this.dgattibutedata.Margin = new System.Windows.Forms.Padding(2);
            this.dgattibutedata.Name = "dgattibutedata";
            this.dgattibutedata.RowHeadersWidth = 51;
            this.dgattibutedata.RowTemplate.Height = 24;
            this.dgattibutedata.Size = new System.Drawing.Size(804, 335);
            this.dgattibutedata.TabIndex = 96;
            this.dgattibutedata.RowHeaderMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgattibutedata_RowHeaderMouseDoubleClick);
            // 
            // TSUpdate
            // 
            this.TSUpdate.HeaderText = "Update";
            this.TSUpdate.Name = "TSUpdate";
            // 
            // AttributeName
            // 
            this.AttributeName.HeaderText = "Attribute";
            this.AttributeName.MinimumWidth = 6;
            this.AttributeName.Name = "AttributeName";
            this.AttributeName.ReadOnly = true;
            this.AttributeName.Width = 200;
            // 
            // AttributeType
            // 
            this.AttributeType.HeaderText = "Type";
            this.AttributeType.MinimumWidth = 6;
            this.AttributeType.Name = "AttributeType";
            this.AttributeType.ReadOnly = true;
            this.AttributeType.Visible = false;
            this.AttributeType.Width = 25;
            // 
            // AttributeValue
            // 
            this.AttributeValue.HeaderText = "Value";
            this.AttributeValue.MinimumWidth = 6;
            this.AttributeValue.Name = "AttributeValue";
            this.AttributeValue.Width = 250;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(256, 8);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 17);
            this.label2.TabIndex = 95;
            this.label2.Text = "Value";
            // 
            // txtattfieldvalue
            // 
            this.txtattfieldvalue.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtattfieldvalue.Location = new System.Drawing.Point(309, 4);
            this.txtattfieldvalue.Margin = new System.Windows.Forms.Padding(2);
            this.txtattfieldvalue.Name = "txtattfieldvalue";
            this.txtattfieldvalue.Size = new System.Drawing.Size(177, 24);
            this.txtattfieldvalue.TabIndex = 94;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(10, 8);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 17);
            this.label1.TabIndex = 93;
            this.label1.Text = "Attribute";
            // 
            // txtattfieldname
            // 
            this.txtattfieldname.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtattfieldname.Location = new System.Drawing.Point(75, 4);
            this.txtattfieldname.Margin = new System.Windows.Forms.Padding(2);
            this.txtattfieldname.Name = "txtattfieldname";
            this.txtattfieldname.Size = new System.Drawing.Size(177, 24);
            this.txtattfieldname.TabIndex = 92;
            // 
            // tbpgview
            // 
            this.tbpgview.Controls.Add(this.tvwpreviewconntype1);
            this.tbpgview.Location = new System.Drawing.Point(4, 22);
            this.tbpgview.Name = "tbpgview";
            this.tbpgview.Padding = new System.Windows.Forms.Padding(3);
            this.tbpgview.Size = new System.Drawing.Size(827, 372);
            this.tbpgview.TabIndex = 0;
            this.tbpgview.Text = "List";
            this.tbpgview.UseVisualStyleBackColor = true;
            this.tbpgview.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // tvwpreviewconntype1
            // 
            this.tvwpreviewconntype1.BackColor = System.Drawing.Color.White;
            this.tvwpreviewconntype1.CheckBoxes = true;
            this.tvwpreviewconntype1.Location = new System.Drawing.Point(3, 3);
            this.tvwpreviewconntype1.Name = "tvwpreviewconntype1";
            this.tvwpreviewconntype1.Size = new System.Drawing.Size(812, 361);
            this.tvwpreviewconntype1.TabIndex = 101;
            this.tvwpreviewconntype1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvwconnection_AfterSelect);
            this.tvwpreviewconntype1.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvwconnection_NodeMouseDoubleClick);
            // 
            // pbar1
            // 
            this.pbar1.Location = new System.Drawing.Point(163, 398);
            this.pbar1.Name = "pbar1";
            this.pbar1.Size = new System.Drawing.Size(547, 11);
            this.pbar1.TabIndex = 104;
            this.pbar1.Visible = false;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(12, 439);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(316, 17);
            this.label3.TabIndex = 111;
            this.label3.Text = "label3";
            // 
            // btntvwtoggle
            // 
            this.btntvwtoggle.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btntvwtoggle.Location = new System.Drawing.Point(1056, 19);
            this.btntvwtoggle.Name = "btntvwtoggle";
            this.btntvwtoggle.Size = new System.Drawing.Size(80, 26);
            this.btntvwtoggle.TabIndex = 103;
            this.btntvwtoggle.Text = "Toggle";
            this.btntvwtoggle.UseVisualStyleBackColor = true;
            this.btntvwtoggle.Click += new System.EventHandler(this.btntvwtoggle_Click);
            // 
            // btnselect
            // 
            this.btnselect.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnselect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnselect.Location = new System.Drawing.Point(843, 228);
            this.btnselect.Name = "btnselect";
            this.btnselect.Size = new System.Drawing.Size(85, 26);
            this.btnselect.TabIndex = 102;
            this.btnselect.Text = "Read Conn.";
            this.btnselect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnselect.UseVisualStyleBackColor = true;
            this.btnselect.Click += new System.EventHandler(this.btnselect_Click);
            // 
            // btnselectconnection
            // 
            this.btnselectconnection.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnselectconnection.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnselectconnection.Location = new System.Drawing.Point(843, 200);
            this.btnselectconnection.Name = "btnselectconnection";
            this.btnselectconnection.Size = new System.Drawing.Size(85, 26);
            this.btnselectconnection.TabIndex = 100;
            this.btnselectconnection.Text = "Select";
            this.btnselectconnection.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnselectconnection.UseVisualStyleBackColor = true;
            this.btnselectconnection.Click += new System.EventHandler(this.btnselectconnection_Click);
            // 
            // rtxtgradechange
            // 
            this.rtxtgradechange.Location = new System.Drawing.Point(1013, 96);
            this.rtxtgradechange.Name = "rtxtgradechange";
            this.rtxtgradechange.Size = new System.Drawing.Size(219, 255);
            this.rtxtgradechange.TabIndex = 107;
            this.rtxtgradechange.Text = "";
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(1013, 63);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(219, 26);
            this.button2.TabIndex = 106;
            this.button2.Text = "A36 to A529-50(L Profile)";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // txtimpfeet
            // 
            this.txtimpfeet.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtimpfeet.Location = new System.Drawing.Point(1287, 373);
            this.txtimpfeet.Name = "txtimpfeet";
            this.txtimpfeet.Size = new System.Drawing.Size(99, 26);
            this.txtimpfeet.TabIndex = 117;
            // 
            // txtimpinch
            // 
            this.txtimpinch.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtimpinch.Location = new System.Drawing.Point(1013, 422);
            this.txtimpinch.Name = "txtimpinch";
            this.txtimpinch.Size = new System.Drawing.Size(99, 26);
            this.txtimpinch.TabIndex = 116;
            this.txtimpinch.TextChanged += new System.EventHandler(this.txtimpinch_TextChanged);
            // 
            // txtmm
            // 
            this.txtmm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtmm.Location = new System.Drawing.Point(1013, 393);
            this.txtmm.Name = "txtmm";
            this.txtmm.Size = new System.Drawing.Size(99, 26);
            this.txtmm.TabIndex = 115;
            this.txtmm.TextChanged += new System.EventHandler(this.txtmm_TextChanged);
            // 
            // btngrouping
            // 
            this.btngrouping.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btngrouping.Location = new System.Drawing.Point(843, 256);
            this.btngrouping.Margin = new System.Windows.Forms.Padding(2);
            this.btngrouping.Name = "btngrouping";
            this.btngrouping.Size = new System.Drawing.Size(85, 26);
            this.btngrouping.TabIndex = 118;
            this.btngrouping.Text = "Group";
            this.btngrouping.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btngrouping.UseVisualStyleBackColor = true;
            this.btngrouping.Click += new System.EventHandler(this.btngrouping_Click);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(1013, 362);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(99, 26);
            this.textBox1.TabIndex = 119;
            // 
            // btnsave
            // 
            this.btnsave.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnsave.Location = new System.Drawing.Point(842, 284);
            this.btnsave.Margin = new System.Windows.Forms.Padding(2);
            this.btnsave.Name = "btnsave";
            this.btnsave.Size = new System.Drawing.Size(85, 26);
            this.btnsave.TabIndex = 120;
            this.btnsave.Text = "Save";
            this.btnsave.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnsave.UseVisualStyleBackColor = true;
            this.btnsave.Click += new System.EventHandler(this.btnsave_Click);
            // 
            // btnload
            // 
            this.btnload.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnload.Location = new System.Drawing.Point(842, 312);
            this.btnload.Margin = new System.Windows.Forms.Padding(2);
            this.btnload.Name = "btnload";
            this.btnload.Size = new System.Drawing.Size(85, 26);
            this.btnload.TabIndex = 121;
            this.btnload.Text = "Load";
            this.btnload.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnload.UseVisualStyleBackColor = true;
            this.btnload.Click += new System.EventHandler(this.btnload_Click);
            // 
            // chkexist
            // 
            this.chkexist.AutoSize = true;
            this.chkexist.Location = new System.Drawing.Point(843, 143);
            this.chkexist.Name = "chkexist";
            this.chkexist.Size = new System.Drawing.Size(78, 17);
            this.chkexist.TabIndex = 122;
            this.chkexist.Text = "Use Exist";
            this.chkexist.UseVisualStyleBackColor = true;
            // 
            // chkonTop
            // 
            this.chkonTop.AutoSize = true;
            this.chkonTop.Location = new System.Drawing.Point(843, 1);
            this.chkonTop.Name = "chkonTop";
            this.chkonTop.Size = new System.Drawing.Size(60, 17);
            this.chkonTop.TabIndex = 123;
            this.chkonTop.Text = "onTop";
            this.chkonTop.UseVisualStyleBackColor = true;
            this.chkonTop.CheckedChanged += new System.EventHandler(this.chkonTop_CheckedChanged);
            // 
            // chkIgnoreConnectionCode
            // 
            this.chkIgnoreConnectionCode.AutoSize = true;
            this.chkIgnoreConnectionCode.Location = new System.Drawing.Point(844, 181);
            this.chkIgnoreConnectionCode.Name = "chkIgnoreConnectionCode";
            this.chkIgnoreConnectionCode.Size = new System.Drawing.Size(94, 17);
            this.chkIgnoreConnectionCode.TabIndex = 124;
            this.chkIgnoreConnectionCode.Text = "Conn. Code";
            this.chkIgnoreConnectionCode.UseVisualStyleBackColor = true;
            // 
            // chkselgrp
            // 
            this.chkselgrp.AutoSize = true;
            this.chkselgrp.Location = new System.Drawing.Point(844, 162);
            this.chkselgrp.Name = "chkselgrp";
            this.chkselgrp.Size = new System.Drawing.Size(65, 17);
            this.chkselgrp.TabIndex = 125;
            this.chkselgrp.Text = "Group.";
            this.chkselgrp.UseVisualStyleBackColor = true;
            // 
            // lblguid
            // 
            this.lblguid.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblguid.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblguid.Location = new System.Drawing.Point(375, 398);
            this.lblguid.Name = "lblguid";
            this.lblguid.Size = new System.Drawing.Size(361, 13);
            this.lblguid.TabIndex = 126;
            // 
            // btngrpguid
            // 
            this.btngrpguid.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btngrpguid.Location = new System.Drawing.Point(843, 340);
            this.btngrpguid.Margin = new System.Windows.Forms.Padding(2);
            this.btngrpguid.Name = "btngrpguid";
            this.btngrpguid.Size = new System.Drawing.Size(85, 26);
            this.btngrpguid.TabIndex = 127;
            this.btngrpguid.Text = "Mast.Guid";
            this.btngrpguid.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btngrpguid.UseVisualStyleBackColor = true;
            this.btngrpguid.Click += new System.EventHandler(this.btngrpguid_Click);
            // 
            // btnclear
            // 
            this.btnclear.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnclear.Location = new System.Drawing.Point(843, 368);
            this.btnclear.Margin = new System.Windows.Forms.Padding(2);
            this.btnclear.Name = "btnclear";
            this.btnclear.Size = new System.Drawing.Size(85, 26);
            this.btnclear.TabIndex = 128;
            this.btnclear.Text = "Clear";
            this.btnclear.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnclear.UseVisualStyleBackColor = true;
            this.btnclear.Click += new System.EventHandler(this.btnclear_Click_1);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ConnectionProcessor.Properties.Resources.esskay;
            this.pictureBox1.Location = new System.Drawing.Point(843, 24);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(85, 113);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 129;
            this.pictureBox1.TabStop = false;
            // 
            // frmconnectionprocess
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(939, 422);
            this.Controls.Add(this.pbar1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnclear);
            this.Controls.Add(this.btngrpguid);
            this.Controls.Add(this.lblguid);
            this.Controls.Add(this.chkselgrp);
            this.Controls.Add(this.chkIgnoreConnectionCode);
            this.Controls.Add(this.chkonTop);
            this.Controls.Add(this.chkexist);
            this.Controls.Add(this.btnload);
            this.Controls.Add(this.btnsave);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btngrouping);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtimpfeet);
            this.Controls.Add(this.txtimpinch);
            this.Controls.Add(this.btntvwtoggle);
            this.Controls.Add(this.txtmm);
            this.Controls.Add(this.btnselect);
            this.Controls.Add(this.rtxtgradechange);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnselectconnection);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.lblxsuser);
            this.Controls.Add(this.lblsbar2);
            this.Controls.Add(this.lblsbar1);
            this.Font = new System.Drawing.Font("Verdana", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmconnectionprocess";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ConnectionProcessor";
            this.Load += new System.EventHandler(this.frmPS_Load);
            this.tabControl1.ResumeLayout(false);
            this.tbpgconn.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvconnectioncompare)).EndInit();
            this.tbpgconnupdate.ResumeLayout(false);
            this.tbpgconnupdate.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgattibutedata)).EndInit();
            this.tbpgview.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lblsbar1;
        private System.Windows.Forms.Label lblsbar2;
        private System.Windows.Forms.Label lblxsuser;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tbpgview;
        private System.Windows.Forms.TabPage tbpgconnupdate;
        private System.Windows.Forms.DataGridView dgattibutedata;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtattfieldvalue;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtattfieldname;
        private System.Windows.Forms.Button btnsetattributes;
        private System.Windows.Forms.Button btngetattributes;
        private System.Windows.Forms.Button btnhlp;
        private System.Windows.Forms.Button btngetallattibuteinfo;
        private Tekla.Structures.Dialog.UIControls.Tree tvwpreviewconntype1;
        private System.Windows.Forms.Button btnselectconnection;
        private System.Windows.Forms.Button btnselect;
        private System.Windows.Forms.Button btntvwtoggle;
        private System.Windows.Forms.RichTextBox rtxtgradechange;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtimpfeet;
        private System.Windows.Forms.TextBox txtimpinch;
        private System.Windows.Forms.TabPage tbpgconn;
        private System.Windows.Forms.DataGridView dgvconnectioncompare;
        private System.Windows.Forms.Button btngrouping;
        private System.Windows.Forms.TextBox txtmm;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnsave;
        private System.Windows.Forms.Button btnload;
        private System.Windows.Forms.CheckBox chkexist;
        private System.Windows.Forms.CheckBox chkonTop;
        private System.Windows.Forms.CheckBox chkIgnoreConnectionCode;
        private Tekla.Structures.Dialog.UIControls.Tree tvwpreviewconn;
        private System.Windows.Forms.CheckBox chkselgrp;
        private System.Windows.Forms.Label lblguid;
        private System.Windows.Forms.Button btngrpguid;
        private System.Windows.Forms.Button btnclear;
        private System.Windows.Forms.DataGridViewCheckBoxColumn TSUpdate;
        private System.Windows.Forms.DataGridViewTextBoxColumn AttributeName;
        private System.Windows.Forms.DataGridViewTextBoxColumn AttributeType;
        private System.Windows.Forms.DataGridViewTextBoxColumn AttributeValue;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnsetallattibuteinfo;
        private System.Windows.Forms.ProgressBar pbar1;
    }
}

