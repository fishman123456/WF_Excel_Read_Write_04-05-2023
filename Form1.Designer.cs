﻿namespace WF_Excel_Read_Write_04_05_2023
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtWrite = new System.Windows.Forms.TextBox();
            this.txtRead = new System.Windows.Forms.TextBox();
            this.btnRead = new System.Windows.Forms.Button();
            this.btnWrite = new System.Windows.Forms.Button();
            this.buttSave = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // txtWrite
            // 
            this.txtWrite.Location = new System.Drawing.Point(16, 12);
            this.txtWrite.Multiline = true;
            this.txtWrite.Name = "txtWrite";
            this.txtWrite.Size = new System.Drawing.Size(100, 261);
            this.txtWrite.TabIndex = 0;
            // 
            // txtRead
            // 
            this.txtRead.Location = new System.Drawing.Point(163, 12);
            this.txtRead.Multiline = true;
            this.txtRead.Name = "txtRead";
            this.txtRead.Size = new System.Drawing.Size(100, 261);
            this.txtRead.TabIndex = 1;
            // 
            // btnRead
            // 
            this.btnRead.Location = new System.Drawing.Point(163, 279);
            this.btnRead.Name = "btnRead";
            this.btnRead.Size = new System.Drawing.Size(100, 23);
            this.btnRead.TabIndex = 2;
            this.btnRead.Text = "Read";
            this.btnRead.UseVisualStyleBackColor = true;
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            // 
            // btnWrite
            // 
            this.btnWrite.Location = new System.Drawing.Point(16, 279);
            this.btnWrite.Name = "btnWrite";
            this.btnWrite.Size = new System.Drawing.Size(100, 23);
            this.btnWrite.TabIndex = 3;
            this.btnWrite.Text = "Write";
            this.btnWrite.UseVisualStyleBackColor = true;
            this.btnWrite.Click += new System.EventHandler(this.btnWrite_Click);
            // 
            // buttSave
            // 
            this.buttSave.Location = new System.Drawing.Point(16, 315);
            this.buttSave.Name = "buttSave";
            this.buttSave.Size = new System.Drawing.Size(247, 23);
            this.buttSave.TabIndex = 4;
            this.buttSave.Text = "SaveAs";
            this.buttSave.UseVisualStyleBackColor = true;
            this.buttSave.Click += new System.EventHandler(this.buttSave_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 361);
            this.Controls.Add(this.buttSave);
            this.Controls.Add(this.btnWrite);
            this.Controls.Add(this.btnRead);
            this.Controls.Add(this.txtRead);
            this.Controls.Add(this.txtWrite);
            this.MaximumSize = new System.Drawing.Size(300, 400);
            this.Name = "Form1";
            this.Text = "U-83";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtWrite;
        private System.Windows.Forms.TextBox txtRead;
        private System.Windows.Forms.Button btnRead;
        private System.Windows.Forms.Button btnWrite;
        private System.Windows.Forms.Button buttSave;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

