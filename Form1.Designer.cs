
namespace WorkExcel
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
            this.button2 = new System.Windows.Forms.Button();
            this.txt_path = new System.Windows.Forms.TextBox();
            this.txt_sheet = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btn_to_lower = new System.Windows.Forms.Button();
            this.txt_export = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_to_upper = new System.Windows.Forms.Button();
            this.btn_email = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(36, 26);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(80, 25);
            this.button1.TabIndex = 0;
            this.button1.Text = "Choose file";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(526, 66);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(80, 24);
            this.button2.TabIndex = 1;
            this.button2.Text = "Upload";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // txt_path
            // 
            this.txt_path.Location = new System.Drawing.Point(122, 29);
            this.txt_path.Name = "txt_path";
            this.txt_path.Size = new System.Drawing.Size(366, 20);
            this.txt_path.TabIndex = 2;
            // 
            // txt_sheet
            // 
            this.txt_sheet.Location = new System.Drawing.Point(122, 69);
            this.txt_sheet.Name = "txt_sheet";
            this.txt_sheet.Size = new System.Drawing.Size(366, 20);
            this.txt_sheet.TabIndex = 3;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(36, 96);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(752, 409);
            this.dataGridView1.TabIndex = 4;
            // 
            // btn_to_lower
            // 
            this.btn_to_lower.Location = new System.Drawing.Point(612, 26);
            this.btn_to_lower.Name = "btn_to_lower";
            this.btn_to_lower.Size = new System.Drawing.Size(80, 25);
            this.btn_to_lower.TabIndex = 5;
            this.btn_to_lower.Text = "To Lower";
            this.btn_to_lower.UseVisualStyleBackColor = true;
            this.btn_to_lower.Click += new System.EventHandler(this.txt_format_Click);
            // 
            // txt_export
            // 
            this.txt_export.Location = new System.Drawing.Point(526, 27);
            this.txt_export.Name = "txt_export";
            this.txt_export.Size = new System.Drawing.Size(80, 24);
            this.txt_export.TabIndex = 6;
            this.txt_export.Text = "Export";
            this.txt_export.UseVisualStyleBackColor = true;
            this.txt_export.Click += new System.EventHandler(this.txt_export_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Type sheet name";
            // 
            // btn_to_upper
            // 
            this.btn_to_upper.Location = new System.Drawing.Point(612, 66);
            this.btn_to_upper.Name = "btn_to_upper";
            this.btn_to_upper.Size = new System.Drawing.Size(80, 25);
            this.btn_to_upper.TabIndex = 8;
            this.btn_to_upper.Text = "To Upper";
            this.btn_to_upper.UseVisualStyleBackColor = true;
            this.btn_to_upper.Click += new System.EventHandler(this.btn_to_upper_Click);
            // 
            // btn_email
            // 
            this.btn_email.Location = new System.Drawing.Point(698, 26);
            this.btn_email.Name = "btn_email";
            this.btn_email.Size = new System.Drawing.Size(80, 25);
            this.btn_email.TabIndex = 9;
            this.btn_email.Text = "Fix Email";
            this.btn_email.UseVisualStyleBackColor = true;
            this.btn_email.Click += new System.EventHandler(this.btn_email_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(818, 517);
            this.Controls.Add(this.btn_email);
            this.Controls.Add(this.btn_to_upper);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txt_export);
            this.Controls.Add(this.btn_to_lower);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.txt_sheet);
            this.Controls.Add(this.txt_path);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Transform";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txt_path;
        private System.Windows.Forms.TextBox txt_sheet;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_to_lower;
        private System.Windows.Forms.Button txt_export;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_to_upper;
        private System.Windows.Forms.Button btn_email;
    }
}

