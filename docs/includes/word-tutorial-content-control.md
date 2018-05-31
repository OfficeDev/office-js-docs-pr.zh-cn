<span data-ttu-id="bea02-101">本教程的这一步是，了解如何在文档中创建格式文本内容控件，以及如何插入和替换控件的内容。</span><span class="sxs-lookup"><span data-stu-id="bea02-101">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span> 

> [!NOTE]
> <span data-ttu-id="bea02-p101">此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="bea02-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

<span data-ttu-id="bea02-104">开始执行本教程的这一步之前，建议通过 Word UI 创建和控制格式文本内容控件，以便熟悉此类控件及其属性。</span><span class="sxs-lookup"><span data-stu-id="bea02-104">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="bea02-105">有关详细信息，请参阅[在 Word 中创建用户填写或打印的表单](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b)。</span><span class="sxs-lookup"><span data-stu-id="bea02-105">For details, see [Create forms that users complete or print in Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

> [!NOTE]
> <span data-ttu-id="bea02-106">虽然可通过 UI 添加到 Word 文档的内容控件有好几种，但目前 Word.js 仅支持格式文本内容控件。</span><span class="sxs-lookup"><span data-stu-id="bea02-106">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>


## <a name="create-a-content-control"></a><span data-ttu-id="bea02-107">创建内容控件</span><span class="sxs-lookup"><span data-stu-id="bea02-107">Create a content control</span></span>

1. <span data-ttu-id="bea02-108">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="bea02-108">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="bea02-109">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="bea02-109">Open the file index.html.</span></span>
3. <span data-ttu-id="bea02-110">在包含 `replace-text` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="bea02-110">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. <span data-ttu-id="bea02-111">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="bea02-111">Open the app.js file.</span></span>

5. <span data-ttu-id="bea02-112">在向 `insert-table` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="bea02-112">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="bea02-113">在 `insertTable` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="bea02-113">Below the `insertTable` function, add the following function:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="bea02-p103">将 `TODO1` 替换为以下代码。注意：</span><span class="sxs-lookup"><span data-stu-id="bea02-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="bea02-116">此代码用于在内容控件中包装短语“Office 365”。</span><span class="sxs-lookup"><span data-stu-id="bea02-116">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="bea02-117">它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="bea02-117">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="bea02-118">属性指定内容控件的可见标题。`ContentControl.title`</span><span class="sxs-lookup"><span data-stu-id="bea02-118">The `ContentControl.title` property specifies the visible title of the content control.</span></span> 
   - <span data-ttu-id="bea02-119">属性指定标记，可用于通过 `ContentControlCollection.getByTag` 方法获取对内容控件的引用，将用于稍后出现的函数。`ContentControl.tag`</span><span class="sxs-lookup"><span data-stu-id="bea02-119">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span> 
   - <span data-ttu-id="bea02-120">属性指定控件的外观。`ContentControl.appearance`</span><span class="sxs-lookup"><span data-stu-id="bea02-120">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="bea02-121">使用值“Tags”表示，控件包装在开始标记和结束标记中，且开始标记包含内容控件标题。</span><span class="sxs-lookup"><span data-stu-id="bea02-121">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="bea02-122">其他可取值包括“BoundingBox”和“None”。</span><span class="sxs-lookup"><span data-stu-id="bea02-122">Other possible values are "BoundingBox" and "None".</span></span>
   - <span data-ttu-id="bea02-123">属性指定标记颜色或边界框的边框。`ContentControl.color`</span><span class="sxs-lookup"><span data-stu-id="bea02-123">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="bea02-124">替换内容控件的内容</span><span class="sxs-lookup"><span data-stu-id="bea02-124">Replace the content of the content control</span></span>

1. <span data-ttu-id="bea02-125">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="bea02-125">Open the file index.html.</span></span>
3. <span data-ttu-id="bea02-126">在包含 `create-content-control` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="bea02-126">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>
    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

4. <span data-ttu-id="bea02-127">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="bea02-127">Open the app.js file.</span></span>

5. <span data-ttu-id="bea02-128">在向 `create-content-control` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="bea02-128">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

6. <span data-ttu-id="bea02-129">在 `createContentControl` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="bea02-129">Below the `createContentControl` function, add the following function:</span></span>

    <span data-ttu-id="bea02-130">\`\`\`js    function replaceContentInControl() {      Word.run(function (context) {</span><span class="sxs-lookup"><span data-stu-id="bea02-130">\`\`\`js    function replaceContentInControl() {      Word.run(function (context) {</span></span>
            
            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    <span data-ttu-id="bea02-131">}</span><span class="sxs-lookup"><span data-stu-id="bea02-131"></span></span>
    ``` 

7. Replace `TODO1` with the following code. 
    > [!NOTE]
    > The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag. We use `getFirst` to get a reference to the desired control.

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="bea02-132">测试加载项</span><span class="sxs-lookup"><span data-stu-id="bea02-132">Test the add-in</span></span>

1. <span data-ttu-id="bea02-133">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl+C 两次，停止正在运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="bea02-133">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="bea02-134">否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。</span><span class="sxs-lookup"><span data-stu-id="bea02-134">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
     > [!NOTE]
     > <span data-ttu-id="bea02-135">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。</span><span class="sxs-lookup"><span data-stu-id="bea02-135">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="bea02-136">为此，需要终止服务器进程，这样才能看到提示并输入生成命令。</span><span class="sxs-lookup"><span data-stu-id="bea02-136">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="bea02-137">生成后，重启服务器。</span><span class="sxs-lookup"><span data-stu-id="bea02-137">After the build, restart the server.</span></span> <span data-ttu-id="bea02-138">接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="bea02-138">The next few steps carry out this process.</span></span>
2. <span data-ttu-id="bea02-139">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="bea02-139">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="bea02-140">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="bea02-140">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="bea02-141">通过关闭任务窗格来重新加载它，再选择“开始”**** 菜单上的“显示任务窗格”****，以重新打开加载项。</span><span class="sxs-lookup"><span data-stu-id="bea02-141">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="bea02-142">在任务窗格中，选择“插入段落”****，以确保文档顶部有包含“Office 365”的段落。</span><span class="sxs-lookup"><span data-stu-id="bea02-142">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>
6. <span data-ttu-id="bea02-143">选择刚刚添加的段落中的短语“Office 365”，再选择“创建内容控件”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="bea02-143">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="bea02-144">观察此短语是否包装在标签为“服务名称”的标记中。</span><span class="sxs-lookup"><span data-stu-id="bea02-144">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>
7. <span data-ttu-id="bea02-145">选择“重命名服务”**** 按钮，并观察内容控件的文本是否变成“Fabrikam Online Productivity Suite”。</span><span class="sxs-lookup"><span data-stu-id="bea02-145">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Word 教程 - 创建内容控件并更改其文本](../images/word-tutorial-content-control.png)
