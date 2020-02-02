namespace Reckon_Connector
{
    partial class frmMain
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
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button6 = new System.Windows.Forms.Button();
            this.cmdGetAPAutofile = new System.Windows.Forms.Button();
            this.cmdSendAPMyob = new System.Windows.Forms.Button();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.button2 = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.button9 = new System.Windows.Forms.Button();
            this.cmdSendARMyob = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.dataGridView4 = new System.Windows.Forms.DataGridView();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.cmdSendContactsAutofile = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.cmdSendAccountAutofile = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.dataGridView5 = new System.Windows.Forms.DataGridView();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.cmdSendClasses = new System.Windows.Forms.Button();
            this.cmdGetClasses = new System.Windows.Forms.Button();
            this.dataGridView6 = new System.Windows.Forms.DataGridView();
            this.cmdSettings = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.RAHRequestTimer = new System.Windows.Forms.Timer(this.components);
            this.cmbSelection = new System.Windows.Forms.ComboBox();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).BeginInit();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.tabPage6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).BeginInit();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage6);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.HotTrack = true;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1066, 740);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage2.Controls.Add(this.button6);
            this.tabPage2.Controls.Add(this.cmdGetAPAutofile);
            this.tabPage2.Controls.Add(this.cmdSendAPMyob);
            this.tabPage2.Controls.Add(this.dataGridView3);
            this.tabPage2.Controls.Add(this.button2);
            this.tabPage2.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage2.ForeColor = System.Drawing.Color.Black;
            this.tabPage2.Location = new System.Drawing.Point(4, 33);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage2.Size = new System.Drawing.Size(1058, 703);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Accounts Payable";
            this.tabPage2.Click += new System.EventHandler(this.tabPage2_Click);
            // 
            // button6
            // 
            this.button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button6.Location = new System.Drawing.Point(786, 17);
            this.button6.Margin = new System.Windows.Forms.Padding(2);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(264, 51);
            this.button6.TabIndex = 8;
            this.button6.Text = "Connect to RAH";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Visible = false;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // cmdGetAPAutofile
            // 
            this.cmdGetAPAutofile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdGetAPAutofile.Location = new System.Drawing.Point(8, 17);
            this.cmdGetAPAutofile.Margin = new System.Windows.Forms.Padding(2);
            this.cmdGetAPAutofile.Name = "cmdGetAPAutofile";
            this.cmdGetAPAutofile.Size = new System.Drawing.Size(237, 51);
            this.cmdGetAPAutofile.TabIndex = 7;
            this.cmdGetAPAutofile.Text = "Get from autofile";
            this.cmdGetAPAutofile.UseVisualStyleBackColor = true;
            this.cmdGetAPAutofile.Click += new System.EventHandler(this.button10_Click);
            // 
            // cmdSendAPMyob
            // 
            this.cmdSendAPMyob.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdSendAPMyob.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdSendAPMyob.Location = new System.Drawing.Point(8, 648);
            this.cmdSendAPMyob.Margin = new System.Windows.Forms.Padding(2);
            this.cmdSendAPMyob.Name = "cmdSendAPMyob";
            this.cmdSendAPMyob.Size = new System.Drawing.Size(237, 51);
            this.cmdSendAPMyob.TabIndex = 6;
            this.cmdSendAPMyob.Text = "Send to Reckon";
            this.cmdSendAPMyob.UseVisualStyleBackColor = true;
            this.cmdSendAPMyob.Click += new System.EventHandler(this.cmdSendAPMyob_Click);
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToAddRows = false;
            this.dataGridView3.AllowUserToDeleteRows = false;
            this.dataGridView3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView3.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Location = new System.Drawing.Point(8, 76);
            this.dataGridView3.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.ReadOnly = true;
            this.dataGridView3.RowTemplate.Height = 33;
            this.dataGridView3.Size = new System.Drawing.Size(1043, 561);
            this.dataGridView3.TabIndex = 3;
            // 
            // button2
            // 
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Location = new System.Drawing.Point(258, 17);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(264, 51);
            this.button2.TabIndex = 0;
            this.button2.Text = "Add New Purchase";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Visible = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage3.Controls.Add(this.button9);
            this.tabPage3.Controls.Add(this.cmdSendARMyob);
            this.tabPage3.Controls.Add(this.button11);
            this.tabPage3.Controls.Add(this.dataGridView4);
            this.tabPage3.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage3.ForeColor = System.Drawing.Color.Black;
            this.tabPage3.Location = new System.Drawing.Point(4, 33);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage3.Size = new System.Drawing.Size(1058, 703);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Accounts Receivable";
            // 
            // button9
            // 
            this.button9.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button9.Location = new System.Drawing.Point(8, 17);
            this.button9.Margin = new System.Windows.Forms.Padding(2);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(237, 51);
            this.button9.TabIndex = 10;
            this.button9.Text = "Get from autofile";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // cmdSendARMyob
            // 
            this.cmdSendARMyob.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdSendARMyob.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdSendARMyob.Location = new System.Drawing.Point(8, 648);
            this.cmdSendARMyob.Margin = new System.Windows.Forms.Padding(2);
            this.cmdSendARMyob.Name = "cmdSendARMyob";
            this.cmdSendARMyob.Size = new System.Drawing.Size(237, 51);
            this.cmdSendARMyob.TabIndex = 9;
            this.cmdSendARMyob.Text = "Send to Reckon";
            this.cmdSendARMyob.UseVisualStyleBackColor = true;
            this.cmdSendARMyob.Click += new System.EventHandler(this.cmdSendARMyob_Click);
            // 
            // button11
            // 
            this.button11.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button11.Location = new System.Drawing.Point(272, 17);
            this.button11.Margin = new System.Windows.Forms.Padding(2);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(263, 51);
            this.button11.TabIndex = 8;
            this.button11.Text = "Add New Sale";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Visible = false;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // dataGridView4
            // 
            this.dataGridView4.AllowUserToAddRows = false;
            this.dataGridView4.AllowUserToDeleteRows = false;
            this.dataGridView4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView4.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView4.Location = new System.Drawing.Point(8, 76);
            this.dataGridView4.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView4.Name = "dataGridView4";
            this.dataGridView4.ReadOnly = true;
            this.dataGridView4.RowTemplate.Height = 33;
            this.dataGridView4.Size = new System.Drawing.Size(1043, 561);
            this.dataGridView4.TabIndex = 6;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.cmdSendContactsAutofile);
            this.tabPage1.Controls.Add(this.button8);
            this.tabPage1.Controls.Add(this.button7);
            this.tabPage1.Controls.Add(this.button4);
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Controls.Add(this.richTextBox3);
            this.tabPage1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage1.ForeColor = System.Drawing.Color.Black;
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage1.Size = new System.Drawing.Size(1058, 703);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Contacts";
            // 
            // cmdSendContactsAutofile
            // 
            this.cmdSendContactsAutofile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdSendContactsAutofile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdSendContactsAutofile.Location = new System.Drawing.Point(8, 648);
            this.cmdSendContactsAutofile.Margin = new System.Windows.Forms.Padding(2);
            this.cmdSendContactsAutofile.Name = "cmdSendContactsAutofile";
            this.cmdSendContactsAutofile.Size = new System.Drawing.Size(237, 51);
            this.cmdSendContactsAutofile.TabIndex = 10;
            this.cmdSendContactsAutofile.Text = "Send to autofile";
            this.cmdSendContactsAutofile.UseVisualStyleBackColor = true;
            this.cmdSendContactsAutofile.Click += new System.EventHandler(this.cmdSendContactsAutofile_Click);
            // 
            // button8
            // 
            this.button8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button8.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button8.Location = new System.Drawing.Point(253, 648);
            this.button8.Margin = new System.Windows.Forms.Padding(2);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(197, 51);
            this.button8.TabIndex = 6;
            this.button8.Text = "Add Customer";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Visible = false;
            // 
            // button7
            // 
            this.button7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button7.Location = new System.Drawing.Point(454, 648);
            this.button7.Margin = new System.Windows.Forms.Padding(2);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(197, 51);
            this.button7.TabIndex = 5;
            this.button7.Text = "Add Supplier";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Visible = false;
            // 
            // button4
            // 
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.Location = new System.Drawing.Point(8, 17);
            this.button4.Margin = new System.Windows.Forms.Padding(2);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(237, 51);
            this.button4.TabIndex = 4;
            this.button4.Text = "Get from Reckon";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(8, 76);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 33;
            this.dataGridView1.Size = new System.Drawing.Size(1043, 561);
            this.dataGridView1.TabIndex = 0;
            // 
            // richTextBox3
            // 
            this.richTextBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.richTextBox3.Location = new System.Drawing.Point(2, 588);
            this.richTextBox3.Margin = new System.Windows.Forms.Padding(2);
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(1054, 113);
            this.richTextBox3.TabIndex = 9;
            this.richTextBox3.Text = "";
            this.richTextBox3.Visible = false;
            // 
            // tabPage4
            // 
            this.tabPage4.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage4.Controls.Add(this.cmdSendAccountAutofile);
            this.tabPage4.Controls.Add(this.button5);
            this.tabPage4.Controls.Add(this.dataGridView2);
            this.tabPage4.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage4.ForeColor = System.Drawing.Color.Black;
            this.tabPage4.Location = new System.Drawing.Point(4, 33);
            this.tabPage4.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage4.Size = new System.Drawing.Size(1058, 703);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Accounts";
            // 
            // cmdSendAccountAutofile
            // 
            this.cmdSendAccountAutofile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdSendAccountAutofile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdSendAccountAutofile.Location = new System.Drawing.Point(8, 648);
            this.cmdSendAccountAutofile.Margin = new System.Windows.Forms.Padding(2);
            this.cmdSendAccountAutofile.Name = "cmdSendAccountAutofile";
            this.cmdSendAccountAutofile.Size = new System.Drawing.Size(237, 51);
            this.cmdSendAccountAutofile.TabIndex = 11;
            this.cmdSendAccountAutofile.Text = "Send to autofile";
            this.cmdSendAccountAutofile.UseVisualStyleBackColor = true;
            this.cmdSendAccountAutofile.Click += new System.EventHandler(this.cmdSendAccountAutofile_Click);
            // 
            // button5
            // 
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.Location = new System.Drawing.Point(8, 17);
            this.button5.Margin = new System.Windows.Forms.Padding(2);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(237, 51);
            this.button5.TabIndex = 3;
            this.button5.Text = "Get from Reckon";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(8, 76);
            this.dataGridView2.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            this.dataGridView2.RowTemplate.Height = 33;
            this.dataGridView2.Size = new System.Drawing.Size(1043, 561);
            this.dataGridView2.TabIndex = 2;
            // 
            // tabPage6
            // 
            this.tabPage6.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage6.Controls.Add(this.button1);
            this.tabPage6.Controls.Add(this.button3);
            this.tabPage6.Controls.Add(this.dataGridView5);
            this.tabPage6.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage6.ForeColor = System.Drawing.Color.Black;
            this.tabPage6.Location = new System.Drawing.Point(4, 33);
            this.tabPage6.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Padding = new System.Windows.Forms.Padding(2);
            this.tabPage6.Size = new System.Drawing.Size(1058, 703);
            this.tabPage6.TabIndex = 5;
            this.tabPage6.Text = "Item Types";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(8, 648);
            this.button1.Margin = new System.Windows.Forms.Padding(2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(237, 51);
            this.button1.TabIndex = 14;
            this.button1.Text = "Send to autofile";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Location = new System.Drawing.Point(8, 17);
            this.button3.Margin = new System.Windows.Forms.Padding(2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(237, 51);
            this.button3.TabIndex = 13;
            this.button3.Text = "Get from Reckon";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // dataGridView5
            // 
            this.dataGridView5.AllowUserToAddRows = false;
            this.dataGridView5.AllowUserToDeleteRows = false;
            this.dataGridView5.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView5.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView5.Location = new System.Drawing.Point(8, 76);
            this.dataGridView5.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView5.Name = "dataGridView5";
            this.dataGridView5.ReadOnly = true;
            this.dataGridView5.RowTemplate.Height = 33;
            this.dataGridView5.Size = new System.Drawing.Size(1043, 561);
            this.dataGridView5.TabIndex = 12;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.cmdSendClasses);
            this.tabPage5.Controls.Add(this.cmdGetClasses);
            this.tabPage5.Controls.Add(this.dataGridView6);
            this.tabPage5.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage5.ForeColor = System.Drawing.Color.Black;
            this.tabPage5.Location = new System.Drawing.Point(4, 33);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(1058, 703);
            this.tabPage5.TabIndex = 6;
            this.tabPage5.Text = "Classes";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // cmdSendClasses
            // 
            this.cmdSendClasses.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cmdSendClasses.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdSendClasses.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdSendClasses.ForeColor = System.Drawing.Color.Black;
            this.cmdSendClasses.Location = new System.Drawing.Point(8, 648);
            this.cmdSendClasses.Margin = new System.Windows.Forms.Padding(2);
            this.cmdSendClasses.Name = "cmdSendClasses";
            this.cmdSendClasses.Size = new System.Drawing.Size(237, 51);
            this.cmdSendClasses.TabIndex = 17;
            this.cmdSendClasses.Text = "Send to autofile";
            this.cmdSendClasses.UseVisualStyleBackColor = true;
            this.cmdSendClasses.Click += new System.EventHandler(this.cmdSendClasses_Click);
            // 
            // cmdGetClasses
            // 
            this.cmdGetClasses.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdGetClasses.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmdGetClasses.ForeColor = System.Drawing.Color.Black;
            this.cmdGetClasses.Location = new System.Drawing.Point(8, 17);
            this.cmdGetClasses.Margin = new System.Windows.Forms.Padding(2);
            this.cmdGetClasses.Name = "cmdGetClasses";
            this.cmdGetClasses.Size = new System.Drawing.Size(237, 51);
            this.cmdGetClasses.TabIndex = 16;
            this.cmdGetClasses.Text = "Get from Reckon";
            this.cmdGetClasses.UseVisualStyleBackColor = true;
            this.cmdGetClasses.Click += new System.EventHandler(this.cmdGetClasses_Click);
            // 
            // dataGridView6
            // 
            this.dataGridView6.AllowUserToAddRows = false;
            this.dataGridView6.AllowUserToDeleteRows = false;
            this.dataGridView6.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView6.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView6.Location = new System.Drawing.Point(8, 76);
            this.dataGridView6.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView6.Name = "dataGridView6";
            this.dataGridView6.ReadOnly = true;
            this.dataGridView6.RowTemplate.Height = 33;
            this.dataGridView6.Size = new System.Drawing.Size(1043, 561);
            this.dataGridView6.TabIndex = 15;
            // 
            // cmdSettings
            // 
            this.cmdSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmdSettings.Location = new System.Drawing.Point(931, 750);
            this.cmdSettings.Margin = new System.Windows.Forms.Padding(2);
            this.cmdSettings.Name = "cmdSettings";
            this.cmdSettings.Size = new System.Drawing.Size(123, 42);
            this.cmdSettings.TabIndex = 10;
            this.cmdSettings.Text = "Settings";
            this.cmdSettings.UseVisualStyleBackColor = true;
            this.cmdSettings.Click += new System.EventHandler(this.cmdSettings_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(12, 759);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(658, 23);
            this.progressBar1.TabIndex = 11;
            this.progressBar1.Visible = false;
            // 
            // RAHRequestTimer
            // 
            this.RAHRequestTimer.Interval = 1000;
            this.RAHRequestTimer.Tick += new System.EventHandler(this.RAHRequestTimer_Tick);
            // 
            // cmbSelection
            // 
            this.cmbSelection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSelection.FormattingEnabled = true;
            this.cmbSelection.Location = new System.Drawing.Point(691, 759);
            this.cmbSelection.Name = "cmbSelection";
            this.cmbSelection.Size = new System.Drawing.Size(220, 26);
            this.cmbSelection.TabIndex = 12;
            this.cmbSelection.SelectedIndexChanged += new System.EventHandler(this.cmbSelection_SelectedIndexChanged);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(1065, 803);
            this.Controls.Add(this.cmbSelection);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.cmdSettings);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.White;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmMain";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Colourworks Reckon Connector Framework v1.0.7";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView4)).EndInit();
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tabPage4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.tabPage6.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView5)).EndInit();
            this.tabPage5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView6)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.RichTextBox richTextBox3;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button cmdSettings;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.Button cmdSendAPMyob;
        private System.Windows.Forms.Button cmdGetAPAutofile;
        private System.Windows.Forms.Button cmdSendContactsAutofile;
        private System.Windows.Forms.Button cmdSendAccountAutofile;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button cmdSendARMyob;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.DataGridView dataGridView4;
        private System.Windows.Forms.TabPage tabPage6;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridView dataGridView5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Timer RAHRequestTimer;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.Button cmdSendClasses;
        private System.Windows.Forms.Button cmdGetClasses;
        private System.Windows.Forms.DataGridView dataGridView6;
        private System.Windows.Forms.ComboBox cmbSelection;
    }
}

