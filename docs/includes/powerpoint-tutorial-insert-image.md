<span data-ttu-id="93ff0-101">本教程的这一步是，检索一天中的[必应](https://www.bing.com)照片，并将图像插入幻灯片。</span><span class="sxs-lookup"><span data-stu-id="93ff0-101">In this step of the tutorial, you'll retrieve the [Bing](https://www.bing.com) photo of the day and insert that image into a slide.</span></span>

> [!NOTE]
> <span data-ttu-id="93ff0-102">此为 PowerPoint 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="93ff0-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="93ff0-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="93ff0-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="add-the-bing-photo-of-the-day-to-a-slide"></a><span data-ttu-id="93ff0-104">向幻灯片添加一天中的必应照片</span><span class="sxs-lookup"><span data-stu-id="93ff0-104">Add the Bing photo of the day to a slide</span></span>

1. <span data-ttu-id="93ff0-105">使用解决方案资源管理器，将 **Controllers** 新文件夹添加到 **HelloWorldWeb** 项目。</span><span class="sxs-lookup"><span data-stu-id="93ff0-105">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![PowerPoint 教程 - 突出显示 HelloWorldWeb 目中 Controllers 文件夹的 Visual Studio 解决方案资源管理器窗口](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="93ff0-107">右键单击“Controllers”**** 文件夹，并依次选择“添加”>“新基架项...”****。</span><span class="sxs-lookup"><span data-stu-id="93ff0-107">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="93ff0-108">在“添加基架”**** 对话框窗口中，依次选择“Web API 2 控制器 - 空”**** 和“添加”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="93ff0-108">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="93ff0-109">在“添加控制器”**** 对话框窗口中，输入“PhotoController”**** 作为控制器名称，再选择“添加”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="93ff0-109">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button.</span></span> <span data-ttu-id="93ff0-110">此时，Visual Studio 创建并打开 **PhotoController.cs** 文件。</span><span class="sxs-lookup"><span data-stu-id="93ff0-110">Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="93ff0-111">将 **PhotoController.cs** 文件的全部内容替换为下列代码，以调用必应服务来检索 Base64 编码字符串形式的一天中照片。</span><span class="sxs-lookup"><span data-stu-id="93ff0-111">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string.</span></span> <span data-ttu-id="93ff0-112">使用 Office JavaScript API 将图像插入文档时，必须将图像数据指定为 Base64 编码字符串。</span><span class="sxs-lookup"><span data-stu-id="93ff0-112">When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

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

6. <span data-ttu-id="93ff0-113">在 **Home.html** 文件中，将 `TODO1` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="93ff0-113">In the **Home.html** file, replace `TODO1` with the following markup.</span></span> <span data-ttu-id="93ff0-114">此标记定义在加载项任务窗格内显示的“插入图像”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="93ff0-114">This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="ms-Button ms-Button--primary" id="insert-image">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Insert Image</span>
        <span class="ms-Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="93ff0-115">在 **Home.js** 文件中，将 `TODO1` 替换为下列代码，以分配“插入图像”**** 按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="93ff0-115">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="93ff0-116">在 **Home.js** 文件中，将 `TODO2` 替换为下列代码，以定义 **insertImage** 函数。</span><span class="sxs-lookup"><span data-stu-id="93ff0-116">In the **Home.js** file, replace `TODO2` with the following code to define the **insertImage** function.</span></span> <span data-ttu-id="93ff0-117">此函数从必应 Web 服务提取图像，再调用 `insertImageFromBase64String` 函数将相应图像插入文档。</span><span class="sxs-lookup"><span data-stu-id="93ff0-117">This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

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

9. <span data-ttu-id="93ff0-118">在 **Home.js** 文件中，将 `TODO3` 替换为下列代码，以定义 `insertImageFromBase64String` 函数。</span><span class="sxs-lookup"><span data-stu-id="93ff0-118">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function.</span></span> <span data-ttu-id="93ff0-119">此函数使用 Office JavaScript API 将图像插入文档。</span><span class="sxs-lookup"><span data-stu-id="93ff0-119">This function uses the Office JavaScript API to insert the image into the document.</span></span> <span data-ttu-id="93ff0-120">注意：</span><span class="sxs-lookup"><span data-stu-id="93ff0-120">Note:</span></span> 

    - <span data-ttu-id="93ff0-121">`coercionType` 选项被指定为 `setSelectedDataAsyc` 请求的第二个参数，指明了要插入的数据的类型。</span><span class="sxs-lookup"><span data-stu-id="93ff0-121">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsyc` request indicates the type of data being inserted.</span></span> 

    - <span data-ttu-id="93ff0-122">`asyncResult` 对象封装 `setSelectedDataAsync` 请求的结果，包括状态和错误消息（如果请求失败的话）。</span><span class="sxs-lookup"><span data-stu-id="93ff0-122">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="93ff0-123">测试加载项</span><span class="sxs-lookup"><span data-stu-id="93ff0-123">Test the add-in</span></span>

1. <span data-ttu-id="93ff0-p107">使用 Visual Studio 的同时，按 `F5` 或选择“开始”**** 按钮启动 PowerPoint，以测试新建的 PowerPoint 加载项，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="93ff0-p107">Using Visual Studio, test the newly created PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="93ff0-127">在 PowerPoint 中，选择功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="93ff0-127">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="93ff0-129">在任务窗格中，选择“插入图像”**** 按钮，将一天中的必应照片添加到当前幻灯片。</span><span class="sxs-lookup"><span data-stu-id="93ff0-129">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![突出显示“插入图像”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="93ff0-131">在 Visual Studio 中，按 `Shift + F5` 或选择“停止”**** 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="93ff0-131">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="93ff0-132">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="93ff0-132">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)