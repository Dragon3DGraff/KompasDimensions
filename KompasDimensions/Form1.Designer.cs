namespace KompasDimensions
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
            this.btnDimensions = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnDimensions
            // 
            this.btnDimensions.Location = new System.Drawing.Point(90, 86);
            this.btnDimensions.Name = "btnDimensions";
            this.btnDimensions.Size = new System.Drawing.Size(234, 66);
            this.btnDimensions.TabIndex = 0;
            this.btnDimensions.Text = "Магическая кнопка";
            this.btnDimensions.UseVisualStyleBackColor = true;
            this.btnDimensions.Click += new System.EventHandler(this.BtnDimensions_Click);
            // 
            // button1
            // 
            this.button1.Image = global::KompasDimensions.Properties.Resources.Weld_L1;
            this.button1.Location = new System.Drawing.Point(363, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(30, 27);
            this.button1.TabIndex = 1;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(405, 249);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnDimensions);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnDimensions;
        private System.Windows.Forms.Button button1;
    }
}

