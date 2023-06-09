using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSPPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office;
using System.IO;
using Microsoft.Office.Core;
using System.Security.Policy;
using System.Drawing;
using static System.Net.WebRequestMethods;

namespace AutoPPT
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("欢迎使用AutoPPT v1.1");
            Console.WriteLine("########################");
            Console.WriteLine("本软件功能为：\n将某个文件夹中所有的图片格式的文件(.jpg;.png;.bmp)合并为一个PPT文件，并储存在您的桌面路径下；");
            Console.WriteLine("##版本更新####################");
            Console.WriteLine("V1.1版本更新：支持用户输入图片边距比例Margin(范围0-80)");
            Console.WriteLine("########################");
            Console.WriteLine("第一步：请您输入图片路径：");
            string targetDir = Console.ReadLine();
            while (!new DirectoryInfo(targetDir).Exists)
            {
                Console.WriteLine("您输入的图片路径有误，请重新");
                targetDir = Console.ReadLine();
            }
            Console.WriteLine("第二步：请您输入图片路径(默认请直接点击回车，默认无页边距)：");
            string marginStr = Console.ReadLine();
            float margin = 0;
            while (!float.TryParse(marginStr,out margin)&& marginStr != "")
            {
                Console.WriteLine("您输入的页边距有误，请重新");
                marginStr = Console.ReadLine();
            }
            margin = Math.Max(0, Math.Min(80, margin));
            var app = new Application();
            app.Visible = MsoTriState.msoCTrue;
            var ppt = app.Presentations.Add();
            var width = ppt.PageSetup.SlideWidth;
            var height = ppt.PageSetup.SlideHeight;

            // change the background color of slide
            ppt.SlideMaster.Background.Fill.ForeColor.RGB = 0;

            
            var figPathList = GetImagePaths(targetDir);
            figPathList = SortPathByCreatedTime(figPathList);
            int currentSlide = 1;

            foreach (string figPath in figPathList)
            {
                var slide = ppt.Slides.Add(currentSlide, PpSlideLayout.ppLayoutBlank);
                Bitmap img = new Bitmap(figPath);
                var w = img.Width;
                var h = img.Height;
                float figwidth = 0;
                float figheight = 0;
                float x_start = 0;
                float y_start = 0;
                if (w / width > h / height)
                {
                    figwidth = width * (1 - margin / 200);
                    figheight = h * width * (1 - margin / 200) / w;
                    y_start = (height - figheight) / 2;
                    x_start = width * margin / 400;
                }
                else
                {
                    figwidth = w * height * (1 - margin / 200) / h;
                    figheight = height * (1 - margin / 200);
                    x_start = (width - figwidth) / 2;
                    y_start = height * margin / 400;
                }
                var shapes = slide.Shapes;
                var shape = shapes.AddPicture(figPath,
                    MsoTriState.msoFalse, MsoTriState.msoTrue, x_start, y_start, figwidth, figheight);
                currentSlide++;
                app.ActiveWindow.View.GotoSlide(slide.SlideIndex);
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            ppt.SaveAs(Path.Combine(desktopPath,"test.ppt"));
            ppt.Close();
            app.Quit();

        }


        public static List<string> GetImagePaths(string folderPath)
        {
            var imagePaths = new List<string>();

            // 获取目标文件夹中所有文件的路径
            var fileNames = Directory.GetFiles(folderPath);
            // 遍历所有文件，查找图片文件路径
            foreach (var fileName in fileNames)
            {
                var extension = Path.GetExtension(fileName).ToLower();

                // 仅添加扩展名为".png"、".jpg"或".bmp"的文件
                if (extension == ".png" || extension == ".jpg" || extension == ".bmp")
                {
                    imagePaths.Add(fileName);
                }
            }
            // 返回所有图片文件路径
            return imagePaths;
        }
        /// <summary>
        // TODO: 如何根据文件名进行排序？不要让11跟在1后面，而跟在10后面
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string> SortPath(List<string> path)
        {
            var result = new List<string> { };

            string lastPath = null;
            foreach (var temp_path in path)
            {
                if (lastPath != null)
                {
                    int thisPathLength = temp_path.Length;
                    for (int i = 0; i < Math.Min(lastPath.Length, thisPathLength); i++)
                    {


                        //if (lastPath[i] != temp_path[i])
                        //{
                        //    if (lastPath[i] > temp_path[i])
                        //    {
                        //        result.Add(temp_path);
                        //        break;
                        //    }
                        //    else
                        //    {
                        //        result.Add(lastPath);
                        //        break;
                        //    }
                        //}
                    }
                }
                else
                {
                    lastPath = temp_path;
                    result.Add(temp_path);
                }
            }


            return result;
        }

        public static List<string> SortPathByCreatedTime(List<string> path)
        {
            var result = path.OrderByDescending(f => (new FileInfo(f)).LastWriteTime).ToList();
            return result;
        }
    }
}
