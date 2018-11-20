<span data-ttu-id="97bd3-101">本教程的这一步是，了解如何在文档中插入图像、HTML 和表格。</span><span class="sxs-lookup"><span data-stu-id="97bd3-101">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

> [!NOTE]
> <span data-ttu-id="97bd3-p101">此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="97bd3-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="insert-an-image"></a><span data-ttu-id="97bd3-104">插入图像</span><span class="sxs-lookup"><span data-stu-id="97bd3-104">Insert an image</span></span>

1. <span data-ttu-id="97bd3-105">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="97bd3-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="97bd3-106">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="97bd3-106">Open the file index.html.</span></span>
3. <span data-ttu-id="97bd3-107">在包含 `replace-text` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="97bd3-107">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="97bd3-108">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="97bd3-108">Open the app.js file.</span></span>

5. <span data-ttu-id="97bd3-109">在文件顶部附近的 use-strict 代码行正下方，添加下面的代码行。</span><span class="sxs-lookup"><span data-stu-id="97bd3-109">Near the top of the file, just below the use-strict line, add the following line.</span></span> <span data-ttu-id="97bd3-110">此代码行导入另一个文件中的变量。</span><span class="sxs-lookup"><span data-stu-id="97bd3-110">This line imports a variable from another file.</span></span> <span data-ttu-id="97bd3-111">此变量是用于编码图像的 Base64 字符串。</span><span class="sxs-lookup"><span data-stu-id="97bd3-111">The variable is a base 64 string that encodes an image.</span></span> <span data-ttu-id="97bd3-112">若要查看已编码字符串，请打开项目根目录中的 base64Image.js 文件。</span><span class="sxs-lookup"><span data-stu-id="97bd3-112">To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="97bd3-113">在向 `replace-text` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="97bd3-113">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="97bd3-114">在 `replaceText` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="97bd3-114">Below the `replaceText` function, add the following function:</span></span>

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

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

8. <span data-ttu-id="97bd3-115">将 `TODO1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="97bd3-115">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="97bd3-116">请注意，此代码行在文档末尾插入 Base64 编码图像。</span><span class="sxs-lookup"><span data-stu-id="97bd3-116">Note that this line inserts the base 64 encoded image at the end of the document.</span></span> <span data-ttu-id="97bd3-117">（`Paragraph` 对象还包含 `insertInlinePictureFromBase64` 方法和其他 `insert*` 方法。</span><span class="sxs-lookup"><span data-stu-id="97bd3-117">(The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods.</span></span> <span data-ttu-id="97bd3-118">有关示例，请参阅下面的 insertHTML 部分。）</span><span class="sxs-lookup"><span data-stu-id="97bd3-118">See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

## <a name="insert-html"></a><span data-ttu-id="97bd3-119">插入 HTML</span><span class="sxs-lookup"><span data-stu-id="97bd3-119">Insert HTML</span></span>

1. <span data-ttu-id="97bd3-120">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="97bd3-120">Open the file index.html.</span></span>
2. <span data-ttu-id="97bd3-121">在包含 `insert-image` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="97bd3-121">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="97bd3-122">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="97bd3-122">Open the app.js file.</span></span>

4. <span data-ttu-id="97bd3-123">在向 `insert-image` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="97bd3-123">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="97bd3-124">在 `insertImage` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="97bd3-124">Below the `insertImage` function, add the following function:</span></span>

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

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

6. <span data-ttu-id="97bd3-p104">将 `TODO1` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="97bd3-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="97bd3-127">第一行代码在文档末尾添加空白段落。</span><span class="sxs-lookup"><span data-stu-id="97bd3-127">The first line adds a blank paragraph to the end of the document.</span></span> 
   - <span data-ttu-id="97bd3-128">第二行代码在段落末尾插入 HTML 字符串；具体而言是两个段落，一个设置使用 Verdana 字体格式，另一个采用 Word 文档的默认样式。</span><span class="sxs-lookup"><span data-stu-id="97bd3-128">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document.</span></span> <span data-ttu-id="97bd3-129">（如前面的 `insertImage` 方法一样，`context.document.body` 对象还包含 `insert*` 方法。）</span><span class="sxs-lookup"><span data-stu-id="97bd3-129">(As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

## <a name="insert-table"></a><span data-ttu-id="97bd3-130">插入表格</span><span class="sxs-lookup"><span data-stu-id="97bd3-130">Insert Table</span></span>

1. <span data-ttu-id="97bd3-131">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="97bd3-131">Open the file index.html.</span></span>
2. <span data-ttu-id="97bd3-132">在包含 `insert-html` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="97bd3-132">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="97bd3-133">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="97bd3-133">Open the app.js file.</span></span>

4. <span data-ttu-id="97bd3-134">在向 `insert-html` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="97bd3-134">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="97bd3-135">在 `insertHTML` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="97bd3-135">Below the `insertHTML` function, add the following function:</span></span>

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

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

6. <span data-ttu-id="97bd3-136">将 `TODO1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="97bd3-136">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="97bd3-137">请注意，此代码行先使用 `ParagraphCollection.getFirst` 方法获取对第一个段落的引用，再使用 `Paragraph.getNext` 方法获取对第二个段落的引用。</span><span class="sxs-lookup"><span data-stu-id="97bd3-137">Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="97bd3-p107">将 `TODO2` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="97bd3-p107">Replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="97bd3-140">`insertTable` 方法的前两个参数指定行数和列数。</span><span class="sxs-lookup"><span data-stu-id="97bd3-140">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>
   - <span data-ttu-id="97bd3-141">第三个参数指定要在哪里插入表格（在此示例中，是在段落后面插入）。</span><span class="sxs-lookup"><span data-stu-id="97bd3-141">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>
   - <span data-ttu-id="97bd3-142">第四个参数是用于设置表格单元格值的二维数组。</span><span class="sxs-lookup"><span data-stu-id="97bd3-142">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>
   - <span data-ttu-id="97bd3-143">虽然表格采用普通的默认样式，但 `insertTable` 方法返回的 `Table` 对象包含多个成员，其中部分成员用于设置表格样式。</span><span class="sxs-lookup"><span data-stu-id="97bd3-143">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="97bd3-144">测试加载项</span><span class="sxs-lookup"><span data-stu-id="97bd3-144">Test the add-in</span></span>


1. <span data-ttu-id="97bd3-145">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl+C 两次，停止正在运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="97bd3-145">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="97bd3-146">否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="97bd3-146">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="97bd3-147">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。</span><span class="sxs-lookup"><span data-stu-id="97bd3-147">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="97bd3-148">为此，需要终止服务器进程，这样才能看到提示并输入生成命令。</span><span class="sxs-lookup"><span data-stu-id="97bd3-148">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="97bd3-149">生成后，重启服务器。</span><span class="sxs-lookup"><span data-stu-id="97bd3-149">After the build, restart the server.</span></span> <span data-ttu-id="97bd3-150">接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="97bd3-150">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="97bd3-151">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="97bd3-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="97bd3-152">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="97bd3-152">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="97bd3-153">通过关闭任务窗格来重新加载它，再选择“开始”\*\*\*\* 菜单上的“显示任务窗格”\*\*\*\*，以重新打开外接程序。</span><span class="sxs-lookup"><span data-stu-id="97bd3-153">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="97bd3-154">在任务窗格中，至少选择“插入段落”\*\*\*\* 三次，以确保文档中有多个段落。</span><span class="sxs-lookup"><span data-stu-id="97bd3-154">In the taskpane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>
6. <span data-ttu-id="97bd3-155">选择“插入图像”\*\*\*\* 按钮，观察图像是否插入在文档末尾。</span><span class="sxs-lookup"><span data-stu-id="97bd3-155">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>
7. <span data-ttu-id="97bd3-156">选择“插入 HTML”\*\*\*\* 按钮，观察是否在文档末尾插入了两个段落，第一个段落使用 Verdana 字体。</span><span class="sxs-lookup"><span data-stu-id="97bd3-156">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>
8. <span data-ttu-id="97bd3-157">选择“插入表格”\*\*\*\* 按钮，观察是否在第二个段落后面插入了表格。</span><span class="sxs-lookup"><span data-stu-id="97bd3-157">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Word 教程 - 插入图像、HTML 和表格](../images/word-tutorial-insert-image-html-table.png)
