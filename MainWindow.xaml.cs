using System.Diagnostics;
using System.Linq;
using System.Windows;
using Xarial.XCad.SolidWorks;
using SolidWorks.Interop.swconst;
using System.IO;

namespace 连接或创建SW程序
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        //后面新建零件、装配体等都需要这个，所以提出来做字段
        private ISwApplication _swApplication;
        public MainWindow()
        {
            InitializeComponent();
        }

        #region 连接至SolidWorks文档
        //连接至目前正在运行的SolidWorks进程
        private void Button_Click_ConnectionToExist(object sender, RoutedEventArgs e)
        {
            var swProcess= Process.GetProcessesByName("SLDWORKS");
            if (!swProcess.Any())
            {
                Blackboard.Text = "无法连接到SolidWorks，新建进程";
                Process.Start("C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\SLDWORKS.exe");
                return;
            }
            _swApplication = SwApplicationFactory.FromProcess(swProcess.First());
            Blackboard.Text = _swApplication.Version.ToString() + " 连接成功." ;  
        }

        /* 如果应用没有打开，可以使用process.start(ProcessPath)方法打开
        Process.Start("C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\sldworks.exe");
        */
    #endregion


    #region 新建各类SolidWorks文档
    //创建一个SolidWorks零件文件
    private void Button_Click_Part(object sender, RoutedEventArgs e)
        {
            NewDocument(swUserPreferenceStringValue_e.swDefaultTemplatePart);
        }
        //创建一个SolidWorks装配体文件
        private void Button_Click_Assmbly(object sender, RoutedEventArgs e)
        {
            NewDocument(swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
        }
        //创建一个SolidWorks工程图文件
        private void Button_Click_Drawing(object sender, RoutedEventArgs e)
        {
            NewDocument(swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
        }
        #endregion 

        #region 新建SolidWorks的方法
                //提取出的新建方法
                //DocmentTemplateType是默认模版类型，swDefaultTemplatePart、swDefaultTemplateAssembly、swDefaultTemplateDrawing三种基本的模版类型
                private void NewDocument(swUserPreferenceStringValue_e DocmentTemplateType)
                {
                    if (_swApplication == null)
                    {
                        Blackboard.Text = "未连接SolidWorks。";
                        return;
                    }
                    else
                    {
                        var DefaultTempelateName = _swApplication.Sw.GetUserPreferenceStringValue((int)DocmentTemplateType);
                        if (!File.Exists(DefaultTempelateName)) { Blackboard.Text = "零件模版文件不存在"; return; }
                        //NewDocument方法需要4个参数，其中需要获取模版的名字，所以需要模版的位置，所以在上面添加获取模版位置的代码
                        _swApplication.Sw.NewDocument(DefaultTempelateName, 0, 297d, 210d);
                    }
                }
        #endregion

        #region 打开已经存在的文档

        int errors = 0;
        int warings = 0;
        //打开零件文档
        private void Button_Click_OpenPart(object sender, RoutedEventArgs e)
        {
            OpenDocument("Part1.sldprt",(int)swDocumentTypes_e.swDocPART);
        }


        //打开装配体文档
        private void Button_Click_OpenAssmbly(object sender, RoutedEventArgs e)
        {
            OpenDocument("Assmbly1.sldasm", (int)swDocumentTypes_e.swDocASSEMBLY);
        }

        //打开工程图文档
        private void Button_Click_OpenDrawing(object sender, RoutedEventArgs e)
        {
            OpenDocument("Drawing1.slddrw", (int)swDocumentTypes_e.swDocDRAWING);
        }

        #region 打开各类型文档的方法
        private void OpenDocument(string DocumentName, int DocumentType)
        {
            //在C#里面使用\表示转意字符，所以需要在前面加上@，让其只表示路径
            string DocumentPath = Path.Combine(Path.GetDirectoryName(typeof(MainWindow).Assembly.Location), "Model", DocumentName);
            if (_swApplication == null)
            {
                Blackboard.Text = "未连接SolidWorks。";
                return;
            }

            if (!File.Exists(DocumentPath))
            {
                Blackboard.Text = "路径可能存在问题。";
                //在SolidWorks窗口里面提出
                _swApplication.ShowMessageBox($"{DocumentPath} 此文件不存在");
                return;
            }
            var PartDocument = _swApplication.Sw.OpenDoc6(DocumentPath, DocumentType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warings);
            if (PartDocument == null)
            {
                _swApplication.ShowMessageBox("${DocumentPath}打开失败，错误代码：{errors}");
                return;
            }
        }
        #endregion 

        #endregion


    }
}
