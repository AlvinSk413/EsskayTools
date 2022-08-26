
namespace att
{
    partial class frmattributeprocessor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmattributeprocessor));
            this.butChangeAttribute = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtstart = new System.Windows.Forms.TextBox();
            this.txtend = new System.Windows.Forms.TextBox();
            this.chklstAttribute = new System.Windows.Forms.CheckedListBox();
            this.txtattributefolder = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.chklstswapattribute = new System.Windows.Forms.CheckedListBox();
            this.txtattreplace = new System.Windows.Forms.TextBox();
            this.txtattfind = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.butphaseadd = new System.Windows.Forms.Button();
            this.butclear = new System.Windows.Forms.Button();
            this.butswapadd = new System.Windows.Forms.Button();
            this.txtphase = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.chkinsidealso = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblsbar = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // butChangeAttribute
            // 
            this.butChangeAttribute.Location = new System.Drawing.Point(9, 209);
            this.butChangeAttribute.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.butChangeAttribute.Name = "butChangeAttribute";
            this.butChangeAttribute.Size = new System.Drawing.Size(154, 36);
            this.butChangeAttribute.TabIndex = 0;
            this.butChangeAttribute.Text = "Change Phase";
            this.butChangeAttribute.UseVisualStyleBackColor = true;
            this.butChangeAttribute.Click += new System.EventHandler(this.butChangeAttribute_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(164, 175);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Start Phase";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(164, 216);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "End Phase";
            // 
            // txtstart
            // 
            this.txtstart.Location = new System.Drawing.Point(296, 171);
            this.txtstart.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.txtstart.Name = "txtstart";
            this.txtstart.Size = new System.Drawing.Size(106, 23);
            this.txtstart.TabIndex = 4;
            this.txtstart.TextChanged += new System.EventHandler(this.txtstart_TextChanged);
            // 
            // txtend
            // 
            this.txtend.Location = new System.Drawing.Point(296, 212);
            this.txtend.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.txtend.Name = "txtend";
            this.txtend.Size = new System.Drawing.Size(106, 23);
            this.txtend.TabIndex = 5;
            this.txtend.TextChanged += new System.EventHandler(this.txtend_TextChanged);
            // 
            // chklstAttribute
            // 
            this.chklstAttribute.FormattingEnabled = true;
            this.chklstAttribute.Location = new System.Drawing.Point(415, 35);
            this.chklstAttribute.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.chklstAttribute.Name = "chklstAttribute";
            this.chklstAttribute.Size = new System.Drawing.Size(159, 202);
            this.chklstAttribute.TabIndex = 6;
            // 
            // txtattributefolder
            // 
            this.txtattributefolder.Location = new System.Drawing.Point(142, 35);
            this.txtattributefolder.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.txtattributefolder.Name = "txtattributefolder";
            this.txtattributefolder.Size = new System.Drawing.Size(262, 23);
            this.txtattributefolder.TabIndex = 8;
            this.txtattributefolder.TextChanged += new System.EventHandler(this.txtattributefolder_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(141, 11);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(105, 17);
            this.label3.TabIndex = 7;
            this.label3.Text = "Attribute Folder";
            // 
            // chklstswapattribute
            // 
            this.chklstswapattribute.FormattingEnabled = true;
            this.chklstswapattribute.Location = new System.Drawing.Point(424, 340);
            this.chklstswapattribute.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.chklstswapattribute.Name = "chklstswapattribute";
            this.chklstswapattribute.Size = new System.Drawing.Size(309, 166);
            this.chklstswapattribute.TabIndex = 9;
            // 
            // txtattreplace
            // 
            this.txtattreplace.Location = new System.Drawing.Point(266, 442);
            this.txtattreplace.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.txtattreplace.Name = "txtattreplace";
            this.txtattreplace.Size = new System.Drawing.Size(205, 23);
            this.txtattreplace.TabIndex = 13;
            // 
            // txtattfind
            // 
            this.txtattfind.Location = new System.Drawing.Point(266, 405);
            this.txtattfind.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.txtattfind.Name = "txtattfind";
            this.txtattfind.Size = new System.Drawing.Size(205, 23);
            this.txtattfind.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(192, 446);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 17);
            this.label4.TabIndex = 11;
            this.label4.Text = "Replace";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(192, 408);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 17);
            this.label5.TabIndex = 10;
            this.label5.Text = "Find";
            // 
            // butphaseadd
            // 
            this.butphaseadd.Location = new System.Drawing.Point(9, 126);
            this.butphaseadd.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.butphaseadd.Name = "butphaseadd";
            this.butphaseadd.Size = new System.Drawing.Size(154, 36);
            this.butphaseadd.TabIndex = 14;
            this.butphaseadd.Text = "Add";
            this.butphaseadd.UseVisualStyleBackColor = true;
            this.butphaseadd.Click += new System.EventHandler(this.butadd_Click);
            // 
            // butclear
            // 
            this.butclear.Location = new System.Drawing.Point(9, 168);
            this.butclear.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.butclear.Name = "butclear";
            this.butclear.Size = new System.Drawing.Size(154, 36);
            this.butclear.TabIndex = 15;
            this.butclear.Text = "`";
            this.butclear.UseVisualStyleBackColor = true;
            this.butclear.Click += new System.EventHandler(this.butclear_Click);
            // 
            // butswapadd
            // 
            this.butswapadd.Location = new System.Drawing.Point(206, 362);
            this.butswapadd.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.butswapadd.Name = "butswapadd";
            this.butswapadd.Size = new System.Drawing.Size(89, 36);
            this.butswapadd.TabIndex = 16;
            this.butswapadd.Text = "+Swap";
            this.butswapadd.UseVisualStyleBackColor = true;
            this.butswapadd.Click += new System.EventHandler(this.butswapadd_Click);
            // 
            // txtphase
            // 
            this.txtphase.Location = new System.Drawing.Point(296, 130);
            this.txtphase.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.txtphase.Name = "txtphase";
            this.txtphase.Size = new System.Drawing.Size(106, 23);
            this.txtphase.TabIndex = 18;
            this.txtphase.TextChanged += new System.EventHandler(this.txtphase_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(164, 134);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 17);
            this.label6.TabIndex = 17;
            this.label6.Text = "Existing Phase";
            // 
            // chkinsidealso
            // 
            this.chkinsidealso.AutoSize = true;
            this.chkinsidealso.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkinsidealso.ForeColor = System.Drawing.Color.Red;
            this.chkinsidealso.Location = new System.Drawing.Point(142, 75);
            this.chkinsidealso.Margin = new System.Windows.Forms.Padding(4);
            this.chkinsidealso.Name = "chkinsidealso";
            this.chkinsidealso.Size = new System.Drawing.Size(232, 21);
            this.chkinsidealso.TabIndex = 20;
            this.chkinsidealso.Text = "Change inside also (Critical)";
            this.chkinsidealso.UseVisualStyleBackColor = true;
            this.chkinsidealso.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(424, 11);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(79, 17);
            this.label7.TabIndex = 21;
            this.label7.Text = "New Phase";
            // 
            // lblsbar
            // 
            this.lblsbar.AutoSize = true;
            this.lblsbar.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.lblsbar.Location = new System.Drawing.Point(13, 250);
            this.lblsbar.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblsbar.Name = "lblsbar";
            this.lblsbar.Size = new System.Drawing.Size(0, 17);
            this.lblsbar.TabIndex = 22;
            // 
            // frmattributeprocessor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(590, 275);
            this.Controls.Add(this.lblsbar);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.chkinsidealso);
            this.Controls.Add(this.txtphase);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.butswapadd);
            this.Controls.Add(this.butclear);
            this.Controls.Add(this.butphaseadd);
            this.Controls.Add(this.txtattreplace);
            this.Controls.Add(this.txtattfind);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.chklstswapattribute);
            this.Controls.Add(this.txtattributefolder);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.chklstAttribute);
            this.Controls.Add(this.txtend);
            this.Controls.Add(this.txtstart);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.butChangeAttribute);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmattributeprocessor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Attribute Changer. Ver1.0";
            this.Load += new System.EventHandler(this.frmattributeprocessor_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button butChangeAttribute;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtstart;
        private System.Windows.Forms.TextBox txtend;
        private System.Windows.Forms.CheckedListBox chklstAttribute;
        private System.Windows.Forms.TextBox txtattributefolder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckedListBox chklstswapattribute;
        private System.Windows.Forms.TextBox txtattreplace;
        private System.Windows.Forms.TextBox txtattfind;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button butphaseadd;
        private System.Windows.Forms.Button butclear;
        private System.Windows.Forms.Button butswapadd;
        private System.Windows.Forms.TextBox txtphase;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox chkinsidealso;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblsbar;
    }
}
