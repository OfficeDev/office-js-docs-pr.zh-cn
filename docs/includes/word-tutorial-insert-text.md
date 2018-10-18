<span data-ttu-id="27d8e-101">本教程的这一步是，先以编程方式测试加载项是否支持用户的当前版本 Word，再在文档中插入段落。</span><span class="sxs-lookup"><span data-stu-id="27d8e-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph in the document.</span></span>

> [!NOTE]
> <span data-ttu-id="27d8e-p101">此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="27d8e-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="27d8e-104">编码加载项</span><span class="sxs-lookup"><span data-stu-id="27d8e-104">Code the add-in</span></span>

1. <span data-ttu-id="27d8e-105">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="27d8e-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="27d8e-106">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="27d8e-106">Open the file index.html.</span></span>
3. <span data-ttu-id="27d8e-107">将 `TODO1` 替换为以下标记：</span><span class="sxs-lookup"><span data-stu-id="27d8e-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="27d8e-108">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="27d8e-108">Open the app.js file.</span></span>
5. <span data-ttu-id="27d8e-109">将 `TODO1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="27d8e-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="27d8e-110">此代码用于确定用户的 Word 版本是否支持包含本教程所有阶段使用的全部 API 的 Word.js 版本。</span><span class="sxs-lookup"><span data-stu-id="27d8e-110">This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial.</span></span> <span data-ttu-id="27d8e-111">在生产加载项中，若要隐藏或禁用调用不受支持的 API 的 UI，请使用条件块的主体。</span><span class="sxs-lookup"><span data-stu-id="27d8e-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="27d8e-112">这样一来，用户仍可以使用 Word 版本支持的加载项部分。</span><span class="sxs-lookup"><span data-stu-id="27d8e-112">This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    } 
    ```

6. <span data-ttu-id="27d8e-113">将 `TODO2` 替换为下面的代码：</span><span class="sxs-lookup"><span data-stu-id="27d8e-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="27d8e-114">将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="27d8e-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="27d8e-115">请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="27d8e-115">Note the following:</span></span>
   - <span data-ttu-id="27d8e-116">Word.js 业务逻辑会添加到传递给 `Word.run` 的函数中。</span><span class="sxs-lookup"><span data-stu-id="27d8e-116">Your Word.js business logic will be added to the function that is passed to `Word.run`.</span></span> <span data-ttu-id="27d8e-117">此逻辑不会立即执行，</span><span class="sxs-lookup"><span data-stu-id="27d8e-117">This logic does not execute immediately.</span></span> <span data-ttu-id="27d8e-118">而是添加到挂起命令队列中。</span><span class="sxs-lookup"><span data-stu-id="27d8e-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="27d8e-119">方法将所有已排入队列的命令都发送到 Word 以供执行。`context.sync`</span><span class="sxs-lookup"><span data-stu-id="27d8e-119">The `context.sync` method sends all queued commands to Word for execution.</span></span>
   - <span data-ttu-id="27d8e-120">后跟 `catch` 块。`Word.run`</span><span class="sxs-lookup"><span data-stu-id="27d8e-120">The `Word.run` is followed by a `catch` block.</span></span> <span data-ttu-id="27d8e-121">这是应始终遵循的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="27d8e-121">This is a best practice that you should always follow.</span></span> 

    ```js
    function insertParagraph() {
        Word.run(function (context) {
            
            // TODO4: Queue commands to insert a paragraph into the document.

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

8. <span data-ttu-id="27d8e-p106">将 `TODO4` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="27d8e-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="27d8e-124">方法的第一个参数是新段落的文本。`insertParagraph`</span><span class="sxs-lookup"><span data-stu-id="27d8e-124">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>
   - <span data-ttu-id="27d8e-125">第二个参数是应在正文中的什么位置插入段落。</span><span class="sxs-lookup"><span data-stu-id="27d8e-125">The second parameter is the location within the body where the paragraph will be inserted.</span></span> <span data-ttu-id="27d8e-126">如果父对象为正文，其他段落插入选项包括“End”和“Replace”。</span><span class="sxs-lookup"><span data-stu-id="27d8e-126">Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span> 

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");   
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="27d8e-127">测试加载项</span><span class="sxs-lookup"><span data-stu-id="27d8e-127">Test the add-in</span></span>

1. <span data-ttu-id="27d8e-128">打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="27d8e-128">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="27d8e-129">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="27d8e-129">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="27d8e-130">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="27d8e-130">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="27d8e-131">通过以下方法之一旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="27d8e-131">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="27d8e-132">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="27d8e-132">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="27d8e-133">Word Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="27d8e-133">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="27d8e-134">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="27d8e-134">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="27d8e-135">在 Word 的“开始”\*\*\*\* 菜单中，选择“显示任务窗格”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="27d8e-135">On the **Home** menu of Word, select **Show Taskpane**.</span></span>
6. <span data-ttu-id="27d8e-136">在任务窗格中，选择“插入段落”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="27d8e-136">In the taskpane, choose **Insert Paragraph**.</span></span>
7. <span data-ttu-id="27d8e-137">在段落中进行一些更改。</span><span class="sxs-lookup"><span data-stu-id="27d8e-137">Make a change in the paragraph.</span></span> 
8. <span data-ttu-id="27d8e-138">再次选择“插入段落”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="27d8e-138">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="27d8e-139">观察新段落是否位于上一段落之上，因为 `insertParagraph` 方法要在文档正文的“开头”插入内容。</span><span class="sxs-lookup"><span data-stu-id="27d8e-139">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the "start" of the document's body.</span></span>

    ![Word 教程 - 插入段落](../images/word-tutorial-insert-paragraph.png)
