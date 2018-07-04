namespace storeLogger
{
    partial class Form1
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
            this.button1 = new System.Windows.Forms.Button();
            this.Element = new System.Windows.Forms.TextBox();
            this.label = new System.Windows.Forms.Label();
            this.wwd = new System.Windows.Forms.Label();
            this.Quantity = new System.Windows.Forms.TextBox();
            this.wdwd = new System.Windows.Forms.Label();
            this.Hazcard = new System.Windows.Forms.TextBox();
            this.ww = new System.Windows.Forms.Label();
            this.Location = new System.Windows.Forms.TextBox();
            this.mn = new System.Windows.Forms.Label();
            this.boughtIn = new System.Windows.Forms.TextBox();
            this.dwd = new System.Windows.Forms.Label();
            this.usedBy = new System.Windows.Forms.TextBox();
            this.wdddd = new System.Windows.Forms.Label();
            this.Stock = new System.Windows.Forms.TextBox();
            this.aa = new System.Windows.Forms.Label();
            this.Company = new System.Windows.Forms.TextBox();
            this.v = new System.Windows.Forms.Label();
            this.Comment = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.rowCount = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button2 = new System.Windows.Forms.Button();
            this.dropDown = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.rowCount)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(401, 281);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(122, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Submit Changes";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Element
            // 
            this.Element.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Element.Location = new System.Drawing.Point(63, 14);
            this.Element.Name = "Element";
            this.Element.Size = new System.Drawing.Size(321, 20);
            this.Element.TabIndex = 1;
            // 
            // label
            // 
            this.label.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label.AutoSize = true;
            this.label.Location = new System.Drawing.Point(12, 17);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(45, 13);
            this.label.TabIndex = 2;
            this.label.Text = "Element";
            // 
            // wwd
            // 
            this.wwd.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wwd.AutoSize = true;
            this.wwd.Location = new System.Drawing.Point(11, 43);
            this.wwd.Name = "wwd";
            this.wwd.Size = new System.Drawing.Size(46, 13);
            this.wwd.TabIndex = 3;
            this.wwd.Text = "Quantity";
            // 
            // Quantity
            // 
            this.Quantity.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Quantity.Location = new System.Drawing.Point(63, 40);
            this.Quantity.Name = "Quantity";
            this.Quantity.Size = new System.Drawing.Size(321, 20);
            this.Quantity.TabIndex = 4;
            // 
            // wdwd
            // 
            this.wdwd.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wdwd.AutoSize = true;
            this.wdwd.Location = new System.Drawing.Point(10, 69);
            this.wdwd.Name = "wdwd";
            this.wdwd.Size = new System.Drawing.Size(47, 13);
            this.wdwd.TabIndex = 5;
            this.wdwd.Text = "Hazcard";
            // 
            // Hazcard
            // 
            this.Hazcard.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Hazcard.Location = new System.Drawing.Point(63, 66);
            this.Hazcard.Name = "Hazcard";
            this.Hazcard.Size = new System.Drawing.Size(321, 20);
            this.Hazcard.TabIndex = 6;
            // 
            // ww
            // 
            this.ww.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ww.AutoSize = true;
            this.ww.Location = new System.Drawing.Point(9, 95);
            this.ww.Name = "ww";
            this.ww.Size = new System.Drawing.Size(48, 13);
            this.ww.TabIndex = 7;
            this.ww.Text = "Location";
            // 
            // Location
            // 
            this.Location.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Location.Location = new System.Drawing.Point(63, 92);
            this.Location.Name = "Location";
            this.Location.Size = new System.Drawing.Size(321, 20);
            this.Location.TabIndex = 8;
            // 
            // mn
            // 
            this.mn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.mn.AutoSize = true;
            this.mn.Location = new System.Drawing.Point(5, 121);
            this.mn.Name = "mn";
            this.mn.Size = new System.Drawing.Size(52, 13);
            this.mn.TabIndex = 9;
            this.mn.Text = "Bought in";
            // 
            // boughtIn
            // 
            this.boughtIn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.boughtIn.Location = new System.Drawing.Point(63, 118);
            this.boughtIn.Name = "boughtIn";
            this.boughtIn.Size = new System.Drawing.Size(321, 20);
            this.boughtIn.TabIndex = 10;
            // 
            // dwd
            // 
            this.dwd.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dwd.AutoSize = true;
            this.dwd.Location = new System.Drawing.Point(11, 147);
            this.dwd.Name = "dwd";
            this.dwd.Size = new System.Drawing.Size(46, 13);
            this.dwd.TabIndex = 11;
            this.dwd.Text = "Used by";
            // 
            // usedBy
            // 
            this.usedBy.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.usedBy.Location = new System.Drawing.Point(63, 144);
            this.usedBy.Name = "usedBy";
            this.usedBy.Size = new System.Drawing.Size(321, 20);
            this.usedBy.TabIndex = 12;
            // 
            // wdddd
            // 
            this.wdddd.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wdddd.AutoSize = true;
            this.wdddd.Location = new System.Drawing.Point(22, 173);
            this.wdddd.Name = "wdddd";
            this.wdddd.Size = new System.Drawing.Size(35, 13);
            this.wdddd.TabIndex = 13;
            this.wdddd.Text = "Stock";
            // 
            // Stock
            // 
            this.Stock.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Stock.Location = new System.Drawing.Point(63, 170);
            this.Stock.Name = "Stock";
            this.Stock.Size = new System.Drawing.Size(321, 20);
            this.Stock.TabIndex = 14;
            // 
            // aa
            // 
            this.aa.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.aa.AutoSize = true;
            this.aa.Location = new System.Drawing.Point(6, 199);
            this.aa.Name = "aa";
            this.aa.Size = new System.Drawing.Size(51, 13);
            this.aa.TabIndex = 15;
            this.aa.Text = "Company";
            // 
            // Company
            // 
            this.Company.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Company.Location = new System.Drawing.Point(63, 196);
            this.Company.Name = "Company";
            this.Company.Size = new System.Drawing.Size(321, 20);
            this.Company.TabIndex = 16;
            // 
            // v
            // 
            this.v.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.v.AutoSize = true;
            this.v.Location = new System.Drawing.Point(6, 225);
            this.v.Name = "v";
            this.v.Size = new System.Drawing.Size(51, 13);
            this.v.TabIndex = 17;
            this.v.Text = "Extra info";
            // 
            // Comment
            // 
            this.Comment.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Comment.Location = new System.Drawing.Point(63, 222);
            this.Comment.Name = "Comment";
            this.Comment.Size = new System.Drawing.Size(321, 20);
            this.Comment.TabIndex = 18;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(399, 17);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(29, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "Row";
            // 
            // rowCount
            // 
            this.rowCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.rowCount.Location = new System.Drawing.Point(434, 15);
            this.rowCount.Maximum = new decimal(new int[] {
            20000,
            0,
            0,
            0});
            this.rowCount.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.rowCount.Name = "rowCount";
            this.rowCount.Size = new System.Drawing.Size(89, 20);
            this.rowCount.TabIndex = 20;
            this.rowCount.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.rowCount.ValueChanged += new System.EventHandler(this.rowCount_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(63, 249);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 13);
            this.label1.TabIndex = 21;
            this.label1.Text = "Current document open:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Excel Documents|*.xlsx";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(273, 281);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(122, 23);
            this.button2.TabIndex = 22;
            this.button2.Text = "Open Excel XML";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dropDown
            // 
            this.dropDown.FormattingEnabled = true;
            this.dropDown.Location = new System.Drawing.Point(401, 69);
            this.dropDown.Name = "dropDown";
            this.dropDown.Size = new System.Drawing.Size(121, 21);
            this.dropDown.TabIndex = 25;
            this.dropDown.SelectedIndexChanged += new System.EventHandler(this.dropDown_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(398, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 26);
            this.label2.TabIndex = 26;
            this.label2.Text = "Select sheet that\r\nyou would like to edit.\r\n";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(535, 313);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dropDown);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rowCount);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.Comment);
            this.Controls.Add(this.v);
            this.Controls.Add(this.Company);
            this.Controls.Add(this.aa);
            this.Controls.Add(this.Stock);
            this.Controls.Add(this.wdddd);
            this.Controls.Add(this.usedBy);
            this.Controls.Add(this.dwd);
            this.Controls.Add(this.boughtIn);
            this.Controls.Add(this.mn);
            this.Controls.Add(this.Location);
            this.Controls.Add(this.ww);
            this.Controls.Add(this.Hazcard);
            this.Controls.Add(this.wdwd);
            this.Controls.Add(this.Quantity);
            this.Controls.Add(this.wwd);
            this.Controls.Add(this.label);
            this.Controls.Add(this.Element);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Excel Filler";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.rowCount)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox Element;
        private System.Windows.Forms.Label label;
        private System.Windows.Forms.Label wwd;
        private System.Windows.Forms.TextBox Quantity;
        private System.Windows.Forms.Label wdwd;
        private System.Windows.Forms.TextBox Hazcard;
        private System.Windows.Forms.Label ww;
        private System.Windows.Forms.TextBox Location;
        private System.Windows.Forms.Label mn;
        private System.Windows.Forms.TextBox boughtIn;
        private System.Windows.Forms.Label dwd;
        private System.Windows.Forms.TextBox usedBy;
        private System.Windows.Forms.Label wdddd;
        private System.Windows.Forms.TextBox Stock;
        private System.Windows.Forms.Label aa;
        private System.Windows.Forms.TextBox Company;
        private System.Windows.Forms.Label v;
        private System.Windows.Forms.TextBox Comment;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.NumericUpDown rowCount;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox dropDown;
        private System.Windows.Forms.Label label2;
    }
}

