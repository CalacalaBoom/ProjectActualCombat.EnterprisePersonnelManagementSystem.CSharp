using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPMS.DataEntity;
using System.IO;
using Sunny.UI;
using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using MSExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace EPMS.ModuleClass
{
    class MyModule
    {
        public static bool FormHelper = false;
        public static string username;//传递登录名
        public static Byte[] imgBytesln;//将图片保存为字节数组
        public static bool mainFormCon = true;


        //将图片转化为字节数组
        public void Read_Image(OpenFileDialog openF, PictureBox MyImage,UITextBox txt)
        {
            openF.Filter = "图片|*.jpg;*.png;*.gif;*.jpeg;*.bmp";
            if (openF.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    MyImage.Image = System.Drawing.Image.FromFile(openF.FileName);
                    txt.Text = Path.GetFullPath(openF.FileName);
                    string strimg = openF.FileName.ToString();
                    FileStream fs = new FileStream(strimg, FileMode.Open, FileAccess.Read);
                    BinaryReader br = new BinaryReader(fs);
                    imgBytesln = br.ReadBytes((int)fs.Length);
                }
                catch
                {
                    UIMessageBox.Show("您选择的图片不能被读取或文件类型不对！");
                }
            }
        }

        //excel合并单元格
        public void Merge(Worksheet excelSheet,int s1, int s2, int e1, int e2)
        {
            MSExcel.Range range = excelSheet.Range[excelSheet.Cells[s1, s2], excelSheet.Cells[e1, e2]];
            range.Merge(true);
        }
    }
}
