<span data-ttu-id="2c600-101">本教程的这一步是，更改文本字体，并向文本应用嵌入样式和自定义样式。</span><span class="sxs-lookup"><span data-stu-id="2c600-101">In this step of the tutorial, you'll change the font of text, and use both built-in and custom styles on the text.</span></span>

> [!NOTE]
> <span data-ttu-id="2c600-p101">此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="2c600-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="2c600-104">向文本应用嵌入样式</span><span class="sxs-lookup"><span data-stu-id="2c600-104">Apply a built-in style to text</span></span>

1. <span data-ttu-id="2c600-105">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="2c600-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="2c600-106">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="2c600-106">Open the file index.html.</span></span>
3. <span data-ttu-id="2c600-107">在包含 `insert-paragraph` 按钮的 `div` 正下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="2c600-107">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="2c600-108">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="2c600-108">Open the app.js file.</span></span>

5. <span data-ttu-id="2c600-109">在向 `insert-paragraph` 按钮分配单击处理程序的代码行正下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="2c600-109">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="2c600-110">在 `insertParagraph` 函数正下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="2c600-110">Just below the `insertParagraph` function, add the following function:</span></span>

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

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

7. <span data-ttu-id="2c600-111">将 `TODO1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="2c600-111">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2c600-112">请注意，此代码向段落应用样式，但也可以向文本区域应用样式。</span><span class="sxs-lookup"><span data-stu-id="2c600-112">Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="2c600-113">向文本应用自定义样式</span><span class="sxs-lookup"><span data-stu-id="2c600-113">Apply a custom style to text</span></span>

1. <span data-ttu-id="2c600-114">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="2c600-114">Open the file index.html.</span></span>
2. <span data-ttu-id="2c600-115">在包含 `apply-style` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="2c600-115">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="2c600-116">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="2c600-116">Open the app.js file.</span></span>

4. <span data-ttu-id="2c600-117">在向 `apply-style` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="2c600-117">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="2c600-118">在 `applyStyle` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="2c600-118">Below the `applyStyle` function, add the following function:</span></span>

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

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

7. <span data-ttu-id="2c600-119">将 `TODO1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="2c600-119">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2c600-120">请注意，此代码应用的自定义样式尚不存在。</span><span class="sxs-lookup"><span data-stu-id="2c600-120">Note that the code applies a custom style that does not exist yet.</span></span> <span data-ttu-id="2c600-121">将在[测试加载项](#test-the-add-in)步骤中创建 **MyCustomStyle** 样式。</span><span class="sxs-lookup"><span data-stu-id="2c600-121">You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a><span data-ttu-id="2c600-122">更改文本字体</span><span class="sxs-lookup"><span data-stu-id="2c600-122">Change the font of text</span></span>

1. <span data-ttu-id="2c600-123">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="2c600-123">Open the file index.html.</span></span>
2. <span data-ttu-id="2c600-124">在包含 `apply-custom-style` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="2c600-124">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="2c600-125">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="2c600-125">Open the app.js file.</span></span>

4. <span data-ttu-id="2c600-126">在向 `apply-custom-style` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="2c600-126">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="2c600-127">在 `applyCustomStyle` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="2c600-127">Below the `applyCustomStyle` function, add the following function:</span></span>

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

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

7. <span data-ttu-id="2c600-128">将 `TODO1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="2c600-128">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="2c600-129">请注意，此代码使用链接到 `Paragraph.getNext` 方法的 `ParagraphCollection.getFirst` 方法，获取对第二个段落的引用。</span><span class="sxs-lookup"><span data-stu-id="2c600-129">Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="2c600-130">测试加载项</span><span class="sxs-lookup"><span data-stu-id="2c600-130">Test the add-in</span></span>

1. <span data-ttu-id="2c600-131">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl+C 两次，停止正在运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="2c600-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="2c600-132">否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。</span><span class="sxs-lookup"><span data-stu-id="2c600-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="2c600-133">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。</span><span class="sxs-lookup"><span data-stu-id="2c600-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="2c600-134">为此，需要终止服务器进程，这样才能看到提示并输入生成命令。</span><span class="sxs-lookup"><span data-stu-id="2c600-134">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="2c600-135">生成后，重启服务器。</span><span class="sxs-lookup"><span data-stu-id="2c600-135">After the build, you restart the server.</span></span> <span data-ttu-id="2c600-136">接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="2c600-136">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="2c600-137">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="2c600-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="2c600-138">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="2c600-138">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="2c600-139">通过关闭任务窗格来重新加载它，再选择“开始”**** 菜单上的“显示任务窗格”****，以重新打开加载项。</span><span class="sxs-lookup"><span data-stu-id="2c600-139">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="2c600-140">请确保文档中至少有三个段落。</span><span class="sxs-lookup"><span data-stu-id="2c600-140">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="2c600-141">可以选择“插入段落”**** 三次。</span><span class="sxs-lookup"><span data-stu-id="2c600-141">You can choose **Insert Paragraph** three times.</span></span> <span data-ttu-id="2c600-142">*仔细检查文档末尾是否没有空白段落。若有，请予以删除。*</span><span class="sxs-lookup"><span data-stu-id="2c600-142">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>
6. <span data-ttu-id="2c600-143">在 Word 中，创建自定义样式“MyCustomStyle”。</span><span class="sxs-lookup"><span data-stu-id="2c600-143">In Word, create a custom style named "MyCustomStyle".</span></span> <span data-ttu-id="2c600-144">其中可以包含所需的任何格式。</span><span class="sxs-lookup"><span data-stu-id="2c600-144">It can have any formatting that you want.</span></span>
7. <span data-ttu-id="2c600-145">选择“应用样式”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="2c600-145">Choose the **Apply Style** button.</span></span> <span data-ttu-id="2c600-146">第一个段落将采用嵌入样式“明显参考”****。</span><span class="sxs-lookup"><span data-stu-id="2c600-146">The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>
8. <span data-ttu-id="2c600-147">选择“应用自定义样式”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="2c600-147">Choose the **Apply Custom Style** button.</span></span> <span data-ttu-id="2c600-148">最后一个段落将采用自定义样式。</span><span class="sxs-lookup"><span data-stu-id="2c600-148">The last paragraph will be styled with your custom style.</span></span> <span data-ttu-id="2c600-149">（如果好像什么都没有发生，很可能是因为最后一个段落是空白段落。</span><span class="sxs-lookup"><span data-stu-id="2c600-149">(If nothing seems to happen, the last paragraph might be blank.</span></span> <span data-ttu-id="2c600-150">如果是这样，请向其中添加某文本。）</span><span class="sxs-lookup"><span data-stu-id="2c600-150">If so, add some text to it.)</span></span>
9. <span data-ttu-id="2c600-151">选择“更改字体”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="2c600-151">Choose the **Change Font** button.</span></span> <span data-ttu-id="2c600-152">第二个段落的字体更改为 18 磅的粗体 Courier New。</span><span class="sxs-lookup"><span data-stu-id="2c600-152">The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Word 教程 - 应用样式和字体](../images/word-tutorial-apply-styles-and-font.png)
