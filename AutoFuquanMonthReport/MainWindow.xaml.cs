using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Aspose.Words.Drawing;
using System.Drawing;
using System.IO;

namespace AutoFuquanMonthReport
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void AutoReport_Click(object sender, RoutedEventArgs e)
        {
            string templateFile = "外观检查报告模板.docx";
            double ImageWidth = 224.25; double ImageHeight = 168.75;
            var doc = new Document(templateFile);
            
            var pictureTable = doc.GetChildNodes(NodeType.Table, true)[4] as Aspose.Words.Tables.Table;
            var builder = new DocumentBuilder(doc);

            builder.MoveTo(pictureTable.Rows[0].Cells[0].FirstParagraph);


            //(暂时用文件名校验)
            //if (!File.Exists($"PicturesOut/{Path.GetFileName(pictureFileName)}"))
            //{
            //    ImageServices.CompressImage($"{pictureFileName}", $"PicturesOut/{Path.GetFileName(pictureFileName)}", CompressImageFlag);    //只取查找到的第1个文件，TODO：UI提示       
            //}
            builder.InsertImage($"PicturesOut/{System.IO.Path.GetFileName("2811")}", RelativeHorizontalPosition.Margin, 0, RelativeVerticalPosition.Margin, 0, ImageWidth, ImageHeight, WrapType.Inline);



        }
    }
}
