本教程的这一步是，检索一天中的[必应](https://www.bing.com)照片，并将图像插入幻灯片。

> [!NOTE]
> 此为 PowerPoint 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="add-the-bing-photo-of-the-day-to-a-slide"></a>向幻灯片添加一天中的必应照片

1. 使用解决方案资源管理器，将 **Controllers** 新文件夹添加到 **HelloWorldWeb** 项目。

    ![PowerPoint 教程 - 突出显示 HelloWorldWeb 目中 Controllers 文件夹的 Visual Studio 解决方案资源管理器窗口](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. 右键单击“Controllers”****文件夹，并依次选择“添加”>“新基架项...”****。

3. 在“添加基架”****对话框窗口中，依次选择“Web API 2 控制器 - 空”****和“添加”****按钮。 

4. 在“添加控制器”****对话框窗口中，输入“PhotoController”****作为控制器名称，再选择“添加”****按钮。 此时，Visual Studio 创建并打开 **PhotoController.cs** 文件。

5. 将 **PhotoController.cs** 文件的全部内容替换为下列代码，以调用必应服务来检索 Base64 编码字符串形式的一天中照片。 使用 Office JavaScript API 将图像插入文档时，必须将图像数据指定为 Base64 编码字符串。

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. 在 **Home.html** 文件中，将 `TODO1` 替换为以下标记。 此标记定义在加载项任务窗格内显示的“插入图像”****按钮。

    ```html
    <button class="ms-Button ms-Button--primary" id="insert-image">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Insert Image</span>
        <span class="ms-Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. 在 **Home.js** 文件中，将 `TODO1` 替换为下列代码，以分配“插入图像”****按钮的事件处理程序。

    ```js
    $('#insert-image').click(insertImage);
    ```

8. 在 **Home.js** 文件中，将 `TODO2` 替换为下列代码，以定义 **insertImage** 函数。 此函数从必应 Web 服务提取图像，再调用 `insertImageFromBase64String` 函数将相应图像插入文档。

    ```js
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. 在 **Home.js** 文件中，将 `TODO3` 替换为下列代码，以定义 `insertImageFromBase64String` 函数。 此函数使用 Office JavaScript API 将图像插入文档。 注意： 

    - `coercionType` 选项被指定为 `setSelectedDataAsyc` 请求的第二个参数，指明了要插入的数据的类型。 

    - `asyncResult` 对象封装 `setSelectedDataAsync` 请求的结果，包括状态和错误消息（如果请求失败的话）。

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a>测试加载项

1. 使用 Visual Studio 的同时，按 `F5` 或选择“开始”****按钮启动 PowerPoint，以测试新建的 PowerPoint 加载项，功能区中显示有“显示任务窗格”****加载项按钮。加载项本地托管在 IIS 上。

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. 在 PowerPoint 中，选择功能区中的“显示任务窗格”****按钮，以打开加载项任务窗格。

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. 在任务窗格中，选择“插入图像”****按钮，将一天中的必应照片添加到当前幻灯片。

    ![突出显示“插入图像”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-image-button.png)

4. 在 Visual Studio 中，按 `Shift + F5` 或选择“停止”****按钮，以停止加载项。 PowerPoint 在加载项停止时自动关闭。

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)