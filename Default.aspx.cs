/**
* ASP实现web上传打印 for Windows
* 使用默认打印机打印 
* 对应的IIS应用池需要授权系统账号
* author @NewFuture
*/
using System;
using System.Diagnostics;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;
using System.Web;
using System.Threading.Tasks;
using System.IO;
using System.Printing;
using System.Linq;

public partial class Print : System.Web.UI.Page
{

    /// <summary>
    /// 日志文件目录
    /// </summary>
    string LOG_DIR
    {
        get { return Server.MapPath("~/file/"); }
    }

    /// <summary>
    /// 文件目录
    /// </summary>
    string FIlE_DIR
    {
        get { return Server.MapPath("~/file/"); }
    }

    /// <summary>
    /// 上传密码
    /// </summary>
    protected string Password = System.Configuration.ConfigurationManager.AppSettings["password"];
    protected bool needPwd { get { return !String.IsNullOrEmpty(this.Password); } }
    /// <summary>
    /// 允许的文件后缀
    /// </summary>
    protected string[] AllowType = { ".pdf", ".doc", ".docx"};
    protected void Page_Load()
    {
        // 新增：显示打印机状态
        /* string status = GetPrinterStatus();
        this.Message.Text = String.Format(
        "<div class='printer-status'>当前打印机状态：{0}</div>", 
        status
        ); */
        if ("POST" == Request.HttpMethod)
        {
            //POST 上传
            Session["msg"] = this.OnPost();
            var url = Request.UrlReferrer;
            if (url == null)
            {
                See_Other(Request.RawUrl);
            }
            else
            {
                See_Other(url.ToString());
            }
        }
        else
        {
            if (!this.needPwd)
            {
                ///密码控件提示
                //this.Message.Text = "<strong>建议优先使用PDF文件</strong>";
            }
            this.Message.Text += Session["msg"];
            Session.Remove("msg");
            return;
        }

    }

    /// <summary>
    /// 处理POST请求的表单
    /// </summary>
    /// <returns></returns>
    protected string OnPost()
    {
        //验证密码
        if (this.needPwd)
        {
            if (this.Password.Trim() != Request.Form["password"])
            {
                return "打印密码无效";
            }
            else
            {
                //暂时保存密码
                Session["pwd"] = Request.Form["password"];
            }
        }

        if (Request.Files.Count < 0)
        {
            return "无文件( ▼-▼ )";
        }

        int copies = 0;
        if (!int.TryParse(Request.Form["copies"], out copies))
        {
            return "份数无效";
        }
        string range = this.GetRange(Request.Form["range"]);
        if (range == null)
        {
            return "打印页码范围[" + Request.Form["range"] + "] 格式错误！";
        }

        //逐个处理文件
        var files = Request.Files;
        string msg = "";
        for (int i = 0; i < files.Count; ++i)
        {
            HttpPostedFile file = files[i];
            if (this.Upload(file, copies, range))
            {
                msg += "<div style='color:green;'>" + file.FileName + "已添加到打印队列</div>";
            }
            else
            {
                msg += "<div style='color:red;'>" + file.FileName + "打印失败</div>";
            }
        }
        return msg;
    }

    /// <summary>
    /// 303重定向
    /// </summary>
    /// <param name="url"></param>
    public void See_Other(string url)
    {
        Response.RedirectLocation = url;
        Response.StatusCode = 303;
        Response.End();
    }

    /// <summary>
    /// 上传打印
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected bool Upload(HttpPostedFile file, int copies, string range)
    {
        string type = GetType(file.FileName);
        if (type == null)
        {
            return false;
        }
        else
        {
            //保存并打印上传文件
            Random r = new Random();
            string path = FIlE_DIR + DateTime.UtcNow.ToString("MM-dd_HH-mm-ss_") + r.Next().ToString() + GetType(file.FileName);
            file.SaveAs(path);
            //打印
            return print(path, copies, range, type);
        }
    }

    /// <summary>
    /// 检查和获取后缀名
    /// </summary>
    /// <param name="filename"></param>
    /// <returns></returns>
    private string GetType(String filename)
    {
        string ext = System.IO.Path.GetExtension(filename).ToLower();
        return Array.IndexOf(AllowType, ext) == -1 ? null : ext;
    }

    private string GetRange(string range)
    {
        if (string.IsNullOrEmpty(range))
        {
            range = "";
        }
        else
        {
            string expr = @"^((\d+(\-\d+)?)\,)*\d+(\-\d+)?$";//^((\d+(\-\d+)?)\,)*\d+(\-\d+)?$
            range = range.Trim().Replace(' ', ',').Replace("，", ",").Trim(',');
            if (!Regex.IsMatch(range, expr))
            {
                return null;
            }
        }
        return range;
    }

    /// <summary>
    /// 打印
    /// </summary>
    /// <param name="path"></param>
    /// <param name="copy"></param>
    /// <param name="range">范围,1-5</param>
    /// <returns></returns>
    protected bool print(string path, int copies = 1, string range = null, string type = ".pdf")
    {
        string cmd, param;
        switch (type)
        {
            case ".pdf":
                //pdf
                // https://www.sumatrapdfreader.org/docs/Command-line-arguments.html
                //cmd = Server.MapPath("~/bin/") + "SumatraPDF.exe";
                //param = string.Format("-print-to-default -silent -print-settings \"{0}x,{1}\"  \"{2}\"", copies, range, path);
                //copies = 1;
                cmd = @"C:\Program Files\LibreOffice\program\soffice.exe";
                param = string.Format("-headless -p \"{0}\"", path);
                break;

            case ".doc":
            case ".docx":
                //word
                //https://superuser.com/questions/352909/how-can-i-get-microsoft-word-2010-to-automatically-quit-after-printing-a-documen
                    cmd = @"C:\Program Files\LibreOffice\program\soffice.exe";
                    param = string.Format("-headless -p \"{0}\"", path);
                break;

            default:
                //word
                cmd = "write";
                param = string.Format("/p \"{0}\"", path);
                break;
        }
        try
        {
            while (copies-- > 0)
            {
                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = cmd,
                        Arguments = param,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };

                    process.OutputDataReceived += (s, e) => log(e.Data, "log");
                    process.ErrorDataReceived += (s, e) => log(e.Data, "error");

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    process.WaitForExit();
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            Console.Write(ex.ToString());
            this.log(ex.ToString(), "exception");

        }
        return false;
    }

    /// <summary>
    /// 获取默认打印机状态
    /// </summary>
    protected string GetPrinterStatus()
    {
        try
        {
            using (LocalPrintServer printServer = new LocalPrintServer())
            {
                PrintQueue defaultQueue = printServer.DefaultPrintQueue;
                if (defaultQueue == null)
                {
                   return "未设置默认打印机";
                }
                // 记录默认打印机详细信息
                log(
                    String.Format("Default Printer: {0}, Task Num: {1}", defaultQueue.Name, defaultQueue.NumberOfJobs),
                    "printer_debug"
                );
                return GetFriendlyStatus(defaultQueue);
            }
        }
        catch (Exception ex)
        {
            log(ex.Message, "printer_error");
            return "状态检测失败：" + ex.Message;
        }
    }
    [System.Web.Services.WebMethod]
    public static string GetPrinterStatusAjax()
    {
        try
        {
            using (LocalPrintServer printServer = new LocalPrintServer())
            {
                PrintQueue defaultQueue = printServer.DefaultPrintQueue;
                if (defaultQueue == null)
                {
                    return "未设置默认打印机";
                }
                return new Print().GetFriendlyStatus(defaultQueue);
            }
        }
        catch (Exception ex)
        {
            return "状态检测失败：" + ex.Message;
        }
    }
    /// <summary>
    /// 将状态代码转换为友好提示
    /// </summary>
    public string GetFriendlyStatus(PrintQueue queue)
    {
        if (queue == null)
        {
            return "未检测到默认打印机";
        }
        // 检查打印队列中的任务数量
        if (queue.NumberOfJobs > 0)
        {
            return String.Format(
                "忙碌（队列中有 {0} 个任务）",
                queue.NumberOfJobs
            );
        }
        // 默认状态
        return queue.QueueStatus == PrintQueueStatus.None ? "就绪" : queue.QueueStatus.ToString();
    }

    /// <summary>
    /// 异步记录日志
    /// </summary>
    /// <param name="msg"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    protected async Task log(string msg, string type = "log")
    {
        if (String.IsNullOrEmpty(msg))
        {
            return;
        }
        string filePath = LOG_DIR + type + ".txt";
        byte[] encodedText = System.Text.Encoding.Unicode.GetBytes(msg + "\n\r");

        using (FileStream sourceStream = new FileStream(filePath,
            FileMode.Append, FileAccess.Write, FileShare.None,
            bufferSize: 4096, useAsync: true))
        {
            await sourceStream.WriteAsync(encodedText, 0, encodedText.Length);
        };
    }
}
