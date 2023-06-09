using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSPPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office;
using Microsoft.Office.Core;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;
using Microsoft.Office.Interop.PowerPoint;
using System.Security.Policy;

namespace PPTStyleChange
{
    public partial class MainForm : Form
    {
        private string pptPath = null;
        public MainForm()
        {
            InitializeComponent();
        }

        private void btn_SelectPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "PowerPoint 文件|*.ppt;*.pptx";
            dialog.Multiselect = false;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string pptFilePath = dialog.FileName;
                if (pptFilePath.Length > 20)
                {
                    
                    this.lbl_pptPath.Text = pptFilePath.Substring(0,3) + "..."+pptFilePath.Substring(pptFilePath.Length - 15,15);
                }
                else
                {
                    this.lbl_pptPath.Text = pptFilePath;
                }
                this.pptPath = pptFilePath;
            }
        }

        private void lbl_dark_Click(object sender, EventArgs e)
        {
            this.panel_dark.BorderStyle = BorderStyle.Fixed3D;

            this.panel_Light.BorderStyle = BorderStyle.None;
            

        }

        private void lbl_light_Click(object sender, EventArgs e)
        {
            this.panel_dark.BorderStyle = BorderStyle.None;
            this.panel_Light.BorderStyle = BorderStyle.Fixed3D;
        }

        private void btn_GO_Click(object sender, EventArgs e)
        {
            int bgColor = 0;
            int txtColor = 0;
            int AllColor = 255 + 255 * 256 + 255 * 256 * 256;
            if (this.panel_dark.BorderStyle == BorderStyle.Fixed3D)
            {
                bgColor = 15 + 15 * 256 + 15 * 256 * 256;
                txtColor = AllColor - bgColor;
            }
            else if (this.panel_Light.BorderStyle == BorderStyle.Fixed3D)
            {
                bgColor = 255 + 255 * 256 + 255 * 256 * 256;
                txtColor = AllColor - bgColor;
            }
            else
            {
                MessageBox.Show("请选择你需要修改的PPT样式");
                return;
            }
            var app = new MSPPT.Application();
            app.Visible = MsoTriState.msoCTrue;
            var ppt = app.Presentations.Open(this.pptPath);
            var width = ppt.PageSetup.SlideWidth;
            var height = ppt.PageSetup.SlideHeight;

            var slides = ppt.Slides;
            ppt.SlideMaster.Background.Fill.ForeColor.RGB = bgColor;
            // 这个是设置PPT的母版背景颜色



            foreach (Slide slide in slides)
            {
                app.ActiveWindow.View.GotoSlide(slide.SlideIndex);

                slide.FollowMasterBackground = MsoTriState.msoTrue;
                slide.Layout = PpSlideLayout.ppLayoutBlank;

                var shapes = slide.Shapes;

                foreach (MSPPT.Shape shape in shapes)
                {
                    var s = shape;

                    var type = s.Fill.Type;



                    int a = 9;
                    //if (s.Fill.Transparency != 0)
                    //{

                    //}
                   

                    //try
                    //{
                    //    //s.Fill.Transparency = 0;
                    //    if (s.Fill.BackColor.Type != MsoColorType.msoColorTypeRGB)
                    //    {
                    //        if (s.Fill.BackColor.RGB < 50 * 256 * 256 | s.Fill.BackColor.RGB > 240 * 256 * 256)
                    //        {
                    //            s.Fill.BackColor.RGB = AllColor - s.Fill.BackColor.RGB;
                    //        }
                                
                    //    }
                    //    if (s.Fill.ForeColor.Type != MsoColorType.msoColorTypeRGB)
                    //    {
                    //        if (s.Fill.ForeColor.RGB < 50*256*256 | s.Fill.ForeColor.RGB >240*256*256)
                    //        {
                    //            s.Fill.ForeColor.RGB = AllColor - s.Fill.ForeColor.RGB;
                    //        }
                    //    }
                        
                    //}
                    //catch
                    //{

                    //}

                    //if (s.HasTextFrame == MsoTriState.msoTrue)
                    //{

                    //    s.TextFrame.TextRange.Font.Color.RGB = txtColor;
                    //}


                }
            }
            //ppt.Save(); 
        }
    }
}
