using Syncfusion.Presentation;
using System;
using System.Windows.Forms;

namespace Convert_PowerPoint_Presentation_to_Image
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
        private Button btnCreate;
        private Label label;
        private void InitializeComponent()
        {
            label = new Label();
            btnCreate = new Button();
            //Label
            label.Location = new System.Drawing.Point(0, 40);
            label.Size = new System.Drawing.Size(426, 35);
            label.Text = "Click the button to Convert PowerPoint to Image.";
            label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            //Button
            btnCreate.Location = new System.Drawing.Point(180, 110);
            btnCreate.Size = new System.Drawing.Size(85, 36);
            btnCreate.Text = "Convert";
            btnCreate.Click += new EventHandler(btnConvert_Click);

            //Create PowerPoint
            ClientSize = new System.Drawing.Size(450, 150);
            Controls.Add(label);
            Controls.Add(btnCreate);
            Text = "Convert PowerPoint to Image";
        }

        #endregion
    }
}

