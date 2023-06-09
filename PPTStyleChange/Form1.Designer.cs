namespace PPTStyleChange
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.lbl_pptPath = new System.Windows.Forms.Label();
            this.btn_SelectPath = new System.Windows.Forms.Button();
            this.panel_dark = new System.Windows.Forms.Panel();
            this.lbl_dark = new System.Windows.Forms.Label();
            this.panel_Light = new System.Windows.Forms.Panel();
            this.lbl_light = new System.Windows.Forms.Label();
            this.btn_GO = new System.Windows.Forms.Button();
            this.panel_dark.SuspendLayout();
            this.panel_Light.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbl_pptPath
            // 
            this.lbl_pptPath.AutoSize = true;
            this.lbl_pptPath.Location = new System.Drawing.Point(27, 29);
            this.lbl_pptPath.Name = "lbl_pptPath";
            this.lbl_pptPath.Size = new System.Drawing.Size(197, 18);
            this.lbl_pptPath.TabIndex = 0;
            this.lbl_pptPath.Text = "请选择你需要修改的PPT";
            // 
            // btn_SelectPath
            // 
            this.btn_SelectPath.Location = new System.Drawing.Point(247, 19);
            this.btn_SelectPath.Name = "btn_SelectPath";
            this.btn_SelectPath.Size = new System.Drawing.Size(75, 39);
            this.btn_SelectPath.TabIndex = 1;
            this.btn_SelectPath.Text = "选择";
            this.btn_SelectPath.UseVisualStyleBackColor = true;
            this.btn_SelectPath.Click += new System.EventHandler(this.btn_SelectPath_Click);
            // 
            // panel_dark
            // 
            this.panel_dark.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.panel_dark.Controls.Add(this.lbl_dark);
            this.panel_dark.Location = new System.Drawing.Point(30, 89);
            this.panel_dark.Name = "panel_dark";
            this.panel_dark.Size = new System.Drawing.Size(139, 120);
            this.panel_dark.TabIndex = 2;
            // 
            // lbl_dark
            // 
            this.lbl_dark.AutoSize = true;
            this.lbl_dark.BackColor = System.Drawing.Color.Transparent;
            this.lbl_dark.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_dark.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.lbl_dark.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.lbl_dark.Location = new System.Drawing.Point(0, 0);
            this.lbl_dark.Margin = new System.Windows.Forms.Padding(15);
            this.lbl_dark.Name = "lbl_dark";
            this.lbl_dark.Padding = new System.Windows.Forms.Padding(38, 50, 38, 50);
            this.lbl_dark.Size = new System.Drawing.Size(138, 118);
            this.lbl_dark.TabIndex = 0;
            this.lbl_dark.Text = "深色调";
            this.lbl_dark.Click += new System.EventHandler(this.lbl_dark_Click);
            // 
            // panel_Light
            // 
            this.panel_Light.BackColor = System.Drawing.Color.White;
            this.panel_Light.Controls.Add(this.lbl_light);
            this.panel_Light.Location = new System.Drawing.Point(183, 89);
            this.panel_Light.Name = "panel_Light";
            this.panel_Light.Size = new System.Drawing.Size(139, 120);
            this.panel_Light.TabIndex = 3;
            // 
            // lbl_light
            // 
            this.lbl_light.AutoSize = true;
            this.lbl_light.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbl_light.Location = new System.Drawing.Point(0, 0);
            this.lbl_light.Margin = new System.Windows.Forms.Padding(15);
            this.lbl_light.Name = "lbl_light";
            this.lbl_light.Padding = new System.Windows.Forms.Padding(38, 50, 38, 50);
            this.lbl_light.Size = new System.Drawing.Size(138, 118);
            this.lbl_light.TabIndex = 1;
            this.lbl_light.Text = "浅色调";
            this.lbl_light.Click += new System.EventHandler(this.lbl_light_Click);
            // 
            // btn_GO
            // 
            this.btn_GO.Location = new System.Drawing.Point(102, 245);
            this.btn_GO.Name = "btn_GO";
            this.btn_GO.Size = new System.Drawing.Size(138, 39);
            this.btn_GO.TabIndex = 4;
            this.btn_GO.Text = "一键换肤";
            this.btn_GO.UseVisualStyleBackColor = true;
            this.btn_GO.Click += new System.EventHandler(this.btn_GO_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(349, 312);
            this.Controls.Add(this.btn_GO);
            this.Controls.Add(this.panel_Light);
            this.Controls.Add(this.panel_dark);
            this.Controls.Add(this.btn_SelectPath);
            this.Controls.Add(this.lbl_pptPath);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PPT一键改色";
            this.panel_dark.ResumeLayout(false);
            this.panel_dark.PerformLayout();
            this.panel_Light.ResumeLayout(false);
            this.panel_Light.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_pptPath;
        private System.Windows.Forms.Button btn_SelectPath;
        private System.Windows.Forms.Panel panel_dark;
        private System.Windows.Forms.Panel panel_Light;
        private System.Windows.Forms.Label lbl_dark;
        private System.Windows.Forms.Label lbl_light;
        private System.Windows.Forms.Button btn_GO;
    }
}

