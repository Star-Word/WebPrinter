<%@ page language="C#" autoeventwireup="true" codefile="Default.aspx.cs" inherits="Print" %>

<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8">
    <title>628线上打印机服务</title>
    <link rel="stylesheet" href="style.css">
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <script src="js/jquery-3.6.0.min.js"></script>
</head>

<body>
    <h1>打印文件</h1>
    <main class="agile-its">
        <h2>628线上打印机服务</h2>
        <div class="print">
            <div class="header">
                <div class="tips">
                    <span>支持格式:</span>
                    <ul class="listtype">
                        <li>PDF(.pdf)[推荐]</li>
                        <!-- <li>图片 (.png .jpg .tiff)</li>
                        <li>文本 (.txt)</li> -->
                        <li>Word (.doc .docx)</li>
                    </ul>
                </div>
                <div class="warn">
					<ul class="listtype">
						<li>仅 PDF 支持<code>"选页"</code>设置</li>
						<li>Word 可能存在排版兼容问题</li>
						<li>仅支持单面打印</li>
						<li>提交后可能需要略微等待文件上传处理</li>
						<li></li>
					</ul>					
				</div>
                <div id="printerStatusContainer">
                    <asp:label id="Message" runat="server"></asp:label>
                    <div id="messages"></div>
                </div>
                <style>
                    .printer-status {
                    padding: 10px;
                    margin: 6px auto 6px;
                    background: #00bcd4;
                    border-radius: 4px;
                    text-align: center;      /* 文本居中 */
                }
                .printer-status[data-status="正常"] { color: green; }
                .printer-status[data-status*="失败"] { color: red; }
                </style>
            </div>

            <form id="upload" method="POST" enctype="multipart/form-data">
                <div class="agileinfo">
                    <div id="filedrag">
                        <span class="uploadtip">点击上传文件<br />
                            或者拖拽至此<br />
                            支持多个文件</span>
                    </div>
                    <input type="file" id="files" name="files[]" multiple="multiple" required="required" accept=".pdf,.doc,.docx" />
                </div>
                <div class="agileinfo inputbox">
                    <% if (needPwd) {
                     string pwd =String.IsNullOrEmpty(Request.Form["password"])?(String)Session["pwd"]:Request.Form["password"].Trim();
                    %>
                    <input type='password' name='password' id='password' onfocus="type='text'" onblur="type='password'" placeholder='打印密码' required='required' value='<%=pwd%>' />
                    <% } %>
                </div>
                <div class="agileinfo inputbox" id="copies">
                    <input name="copies" type="number" value="1" title="份数" placeholder="设置份数" required />
                </div>
                <div class="agileinfo inputbox">
                    <input name="range" type="text" title="页码范围如:2-8 或1,3,5" placeholder="PDF页码:2-5或3,5 (默认所有页)" />
                </div>
                <button type="submit" id="submit" onsubmit="this.disabled=true" disabled>提交</button>
            </form>
        </div>
    </main>
    <footer>
        <p><strong>&lsaquo;&rsaquo;</strong> with <strong>&hearts;</strong> by New Future | <a href="https://github.com/NewFuture/WebPrint">获取源码</a></p>
        <p>部署于2025.05.23</p>
    </footer>
    <script>
    // 每隔2秒更新打印机状态
        setInterval(function() {
            $.ajax({
                type: "POST",
                url: "Default.aspx/GetPrinterStatusAjax",
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function(response) {
                    $("#Message").html("<div class='printer-status'>当前打印机状态：" + response.d + "</div>");
                },
                error: function(xhr, status, error) {
                    console.error("更新失败: " + error);
                }
            });
        }, 2000); // 2000毫秒 = 2秒
    </script>
    <script src="file.js"></script>

</body>

</html>
