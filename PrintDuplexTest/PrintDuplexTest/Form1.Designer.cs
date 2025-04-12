namespace PrintDuplexTest
{
    partial class Form1
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
            printButton = new Button();
            comboBox1 = new ComboBox();
            SuspendLayout();
            // 
            // printButton
            // 
            printButton.Location = new Point(375, 159);
            printButton.Name = "printButton";
            printButton.Size = new Size(94, 29);
            printButton.TabIndex = 0;
            printButton.Text = "button1";
            printButton.UseVisualStyleBackColor = true;
            printButton.Click += printButton_Click;
            // 
            // comboBox1
            // 
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(131, 62);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(348, 28);
            comboBox1.TabIndex = 1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(538, 210);
            Controls.Add(comboBox1);
            Controls.Add(printButton);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ResumeLayout(false);
        }

        #endregion

        private Button printButton;
        private ComboBox comboBox1;
    }
}
