namespace ExtractFuntion
{
    partial class MainForm
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
            this.button_FindSolution = new System.Windows.Forms.Button();
            this.textBox_SolutionPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar_Progress = new System.Windows.Forms.ProgressBar();
            this.label_ProgressPercent = new System.Windows.Forms.Label();
            this.button_StartExtract = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_FindSolution
            // 
            this.button_FindSolution.Location = new System.Drawing.Point(396, 41);
            this.button_FindSolution.Name = "button_FindSolution";
            this.button_FindSolution.Size = new System.Drawing.Size(75, 23);
            this.button_FindSolution.TabIndex = 0;
            this.button_FindSolution.Text = "찾아보기";
            this.button_FindSolution.UseVisualStyleBackColor = true;
            this.button_FindSolution.Click += new System.EventHandler(this.button_FindSolution_Click);
            // 
            // textBox_SolutionPath
            // 
            this.textBox_SolutionPath.Location = new System.Drawing.Point(14, 41);
            this.textBox_SolutionPath.Name = "textBox_SolutionPath";
            this.textBox_SolutionPath.Size = new System.Drawing.Size(376, 21);
            this.textBox_SolutionPath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "솔루션 (프로젝트) 경로";
            // 
            // progressBar_Progress
            // 
            this.progressBar_Progress.Location = new System.Drawing.Point(14, 103);
            this.progressBar_Progress.Name = "progressBar_Progress";
            this.progressBar_Progress.Size = new System.Drawing.Size(376, 23);
            this.progressBar_Progress.TabIndex = 3;
            // 
            // label_ProgressPercent
            // 
            this.label_ProgressPercent.AutoSize = true;
            this.label_ProgressPercent.Location = new System.Drawing.Point(13, 85);
            this.label_ProgressPercent.Name = "label_ProgressPercent";
            this.label_ProgressPercent.Size = new System.Drawing.Size(21, 12);
            this.label_ProgressPercent.TabIndex = 4;
            this.label_ProgressPercent.Text = "0%";
            // 
            // button_StartExtract
            // 
            this.button_StartExtract.BackColor = System.Drawing.Color.LightSteelBlue;
            this.button_StartExtract.Location = new System.Drawing.Point(396, 98);
            this.button_StartExtract.Name = "button_StartExtract";
            this.button_StartExtract.Size = new System.Drawing.Size(75, 32);
            this.button_StartExtract.TabIndex = 5;
            this.button_StartExtract.Text = "추출 시작";
            this.button_StartExtract.UseVisualStyleBackColor = false;
            this.button_StartExtract.Click += new System.EventHandler(this.button_StartExtract_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(492, 147);
            this.Controls.Add(this.button_StartExtract);
            this.Controls.Add(this.label_ProgressPercent);
            this.Controls.Add(this.progressBar_Progress);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_SolutionPath);
            this.Controls.Add(this.button_FindSolution);
            this.Name = "MainForm";
            this.Text = "함수 추출기";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_FindSolution;
        private System.Windows.Forms.TextBox textBox_SolutionPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar_Progress;
        private System.Windows.Forms.Label label_ProgressPercent;
        private System.Windows.Forms.Button button_StartExtract;
    }
}