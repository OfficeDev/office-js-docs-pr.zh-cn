<span data-ttu-id="ef3b4-101">本教程的这一步是，在选定文本区域内外添加文本，并替换选定区域的文本。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-101">In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.</span></span> 

> [!NOTE]
> <span data-ttu-id="ef3b4-p101">此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="add-text-inside-a-range"></a><span data-ttu-id="ef3b4-104">在区域内添加文本</span><span class="sxs-lookup"><span data-stu-id="ef3b4-104">Add text inside a range</span></span>

1. <span data-ttu-id="ef3b4-105">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="ef3b4-106">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-106">Open the file index.html.</span></span>
3. <span data-ttu-id="ef3b4-107">在包含 `change-font` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-107">Below the `div` that contains the `change-font` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>            
    </div>
    ```

4. <span data-ttu-id="ef3b4-108">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-108">Open the app.js file.</span></span>

5. <span data-ttu-id="ef3b4-109">在向 `change-font` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-109">Below the line that assigns a click handler to the `change-font` button, add the following code:</span></span>

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. <span data-ttu-id="ef3b4-110">在 `changeFont` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-110">Below the `changeFont` function, add the following function:</span></span>

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the 
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

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

7. <span data-ttu-id="ef3b4-p102">将 `TODO1` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-p102">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="ef3b4-113">此方法用于在“即点即用”文本区域末尾插入缩写 ["(C2R)"]。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-113">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run".</span></span> <span data-ttu-id="ef3b4-114">它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-114">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="ef3b4-115">方法的第一个参数是要插入到 `Range` 对象的字符串。`Range.insertText`</span><span class="sxs-lookup"><span data-stu-id="ef3b4-115">The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.</span></span>
   - <span data-ttu-id="ef3b4-116">第二个参数指定了应在区域中的什么位置插入其他文本。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-116">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="ef3b4-117">除了“End”外，其他可用选项包括“Start”、“Before”、“After”和“Replace”。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-117">Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span></span> 
   - <span data-ttu-id="ef3b4-118">“End”和“After”的区别在于，“End”在现有区域末尾插入新文本，而“After”则是新建包含字符串的区域，并在现有区域后面插入新区域。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-118">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range.</span></span> <span data-ttu-id="ef3b4-119">同样，“Start”是在现有区域的开头位置插入文本，而“Before”插入的是新区域。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-119">Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range.</span></span> <span data-ttu-id="ef3b4-120">“Replace”将现有区域文本替换为第一个参数中的字符串。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-120">"Replace" replaces the text of the existing range with the string in the first parameter.</span></span>
   - <span data-ttu-id="ef3b4-121">在本教程之前阶段步骤中，正文对象的 insert\* 方法没有“Before”和“After”选项。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-121">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options.</span></span> <span data-ttu-id="ef3b4-122">这是因为不能将内容置于文档正文外。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-122">This is because you can't put content outside of the document's body.</span></span>

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ``` 

8. <span data-ttu-id="ef3b4-123">在下一部分前，将跳过 `TODO2`。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-123">We'll skip over `TODO2` until the next section.</span></span> <span data-ttu-id="ef3b4-124">将 `TODO3` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-124">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="ef3b4-125">此代码类似于在本教程第一阶段中创建的代码，区别在于现在是要在文档末尾（而不是开头）插入新段落。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-125">This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start.</span></span> <span data-ttu-id="ef3b4-126">这一新段落将说明，新文本现属于原始区域。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-126">This new paragraph will demonstrate that the new text is now part of the original range.</span></span>
 
    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text,
                             "End");
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="ef3b4-127">添加代码以将文档属性提取到任务窗格的脚本对象</span><span class="sxs-lookup"><span data-stu-id="ef3b4-127">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="ef3b4-128">在本系列教程前面的所有函数中，都是将命令排入队列，以对 Office 文档执行*写入*操作。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-128">In all the previous functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="ef3b4-129">每个函数结束时都会调用 `context.sync()` 方法，从而将排入队列的命令发送到文档，以供执行。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-129">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="ef3b4-130">不过，在上一步中添加的代码调用的是 `originalRange.text` 属性，这与之前编写的函数明显不同，因为 `originalRange` 对象只是任务窗格脚本中的代理对象。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-130">But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="ef3b4-131">由于它并不了解文档中区域的实际文本，因此它的 `text` 属性无法有实值。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-131">It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value.</span></span> <span data-ttu-id="ef3b4-132">有必要先从文档中提取区域的文本值，再用它设置 `originalRange.text` 的值。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-132">It is necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`.</span></span> <span data-ttu-id="ef3b4-133">只有这样才能调用 `originalRange.text`，而又不会导致异常抛出。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-133">Only then can `originalRange.text` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="ef3b4-134">此提取过程分为三步：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-134">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="ef3b4-135">将命令排入队列，以加载（即提取）代码需要读取的属性。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-135">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>
   2. <span data-ttu-id="ef3b4-136">调用上下文对象的 `sync`方法，从而向文档发送已排入队列的命令以供执行，并返回请求获取的信息。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-136">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>
   3. <span data-ttu-id="ef3b4-137">由于 `sync` 是异步方法，因此请先确保它已完成，然后代码才能调用已提取的属性。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-137">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="ef3b4-138">只要代码需要从 Office 文档*读取*信息，就必须完成这些步骤。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-138">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="ef3b4-139">将 `TODO2` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-139">Replace `TODO2` with the following code.</span></span>
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.
    
            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

2. <span data-ttu-id="ef3b4-p109">由于不能在同一取消分支代码路径中有两个 `return` 语句，因此请删除 `Word.run` 末尾的最后一行代码 `return context.sync();`。本教程稍后将添加最后一个新 `context.sync` 语句。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-p109">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`. You'll add a new final `context.sync` later in this tutorial.</span></span> 
3. <span data-ttu-id="ef3b4-142">剪切并粘贴 `doc.body.insertParagraph` 代码行，以替代 `TODO4`。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-142">Cut the `doc.body.insertParagraph` line and paste in place of `TODO4`.</span></span> 
4. <span data-ttu-id="ef3b4-p110">将 `TODO5` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-p110">Replace `TODO5` with the following code. Note:</span></span>
   - <span data-ttu-id="ef3b4-145">将 `sync` 方法传递到 `then` 函数可确保它不会在 `insertParagraph` 逻辑已排入队列前运行。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-145">Passing the `sync` method to a `then` function ensures that it does not run until the `insertParagraph` logic has been queued.</span></span>
   - <span data-ttu-id="ef3b4-146">由于 `then` 方法调用传递给它的任何函数，并且也不想调用 `sync` 两次，因此请从 context.sync 末尾省略掉“()”。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-146">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of context.sync.</span></span>

    ```js
    .then(context.sync);
    ```

<span data-ttu-id="ef3b4-147">完成后，整个函数应如下所示：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-147">When you are done, the entire function should look like the following:</span></span>

  
```js
function insertTextIntoRange() {
    Word.run(function (context) {
        
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {        
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");            
                }
            )
            .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
``` 

## <a name="add-text-between-ranges"></a><span data-ttu-id="ef3b4-148">在区域间添加文本</span><span class="sxs-lookup"><span data-stu-id="ef3b4-148">Add text between ranges</span></span>

1. <span data-ttu-id="ef3b4-149">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-149">Open the file index.html.</span></span>
2. <span data-ttu-id="ef3b4-150">在包含 `insert-text-into-range` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-150">Below the `div` that contains the `insert-text-into-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>            
    </div>
    ```

3. <span data-ttu-id="ef3b4-151">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-151">Open the app.js file.</span></span>

4. <span data-ttu-id="ef3b4-152">在向 `insert-text-into-range` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-152">Below the line that assigns a click handler to the `insert-text-into-range` button, add the following code:</span></span>

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. <span data-ttu-id="ef3b4-153">在 `insertTextIntoRange` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-153">Below the `insertTextIntoRange` function, add the following function:</span></span>

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert a new range before the 
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the 
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="ef3b4-p111">将 `TODO1` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-p111">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="ef3b4-156">此方法用于在文本为“Office 365”的区域前添加文本为“Office 2019”的区域。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-156">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365".</span></span> <span data-ttu-id="ef3b4-157">它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-157">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="ef3b4-158">方法的第一个参数是要添加的字符串。`Range.insertText`</span><span class="sxs-lookup"><span data-stu-id="ef3b4-158">The first parameter of the `Range.insertText` method is the string to add.</span></span>
   - <span data-ttu-id="ef3b4-159">第二个参数指定了应在区域中的什么位置插入其他文本。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-159">The second parameter specifies where in the range the additional text should be inserted.</span></span> <span data-ttu-id="ef3b4-160">若要详细了解位置选项，请参阅前面介绍的 `insertTextIntoRange` 函数。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-160">For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span></span>

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ``` 

7. <span data-ttu-id="ef3b4-161">将 `TODO2` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-161">Replace `TODO2` with the following code.</span></span> 
 
     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.
    
                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

8. <span data-ttu-id="ef3b4-162">将 `TODO3` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-162">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="ef3b4-163">这一新段落将说明，新文本***不***属于原始选定区域。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-163">This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range.</span></span> <span data-ttu-id="ef3b4-164">原始区域中的文本仍与用户选择它时一样。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-164">The original range still has only the text it had when it was selected.</span></span>
 
    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ``` 

9. <span data-ttu-id="ef3b4-165">将 `TODO4` 替换为下面的代码：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-165">Replace `TODO4` with the following code:</span></span>

    ```js
    .then(context.sync);
    ```


## <a name="replace-the-text-of-a-range"></a><span data-ttu-id="ef3b4-166">替换区域文本</span><span class="sxs-lookup"><span data-stu-id="ef3b4-166">Replace the text of a range</span></span>

1. <span data-ttu-id="ef3b4-167">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-167">Open the file index.html.</span></span>
2. <span data-ttu-id="ef3b4-168">在包含 `insert-text-outside-range` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-168">Below the `div` that contains the `insert-text-outside-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>            
    </div>
    ```

3. <span data-ttu-id="ef3b4-169">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-169">Open the app.js file.</span></span>

4. <span data-ttu-id="ef3b4-170">在向 `insert-text-outside-range` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-170">Below the line that assigns a click handler to the `insert-text-outside-range` button, add the following code:</span></span>

    ```js
    $('#replace-text').click(replaceText);
    ```

5. <span data-ttu-id="ef3b4-171">在 `insertTextBeforeRange` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="ef3b4-171">Below the `insertTextBeforeRange` function, add the following function:</span></span>

    ```js
    function replaceText() {
        Word.run(function (context) {
             
            // TODO1: Queue commands to replace the text.

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

6. <span data-ttu-id="ef3b4-172">将 `TODO1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-172">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="ef3b4-173">请注意，此方法用于将字符串“几个”替换为字符串“许多”。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-173">Note that the method is intended to replace the string "several" with the string "many".</span></span> <span data-ttu-id="ef3b4-174">它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-174">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace"); 
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="ef3b4-175">测试加载项</span><span class="sxs-lookup"><span data-stu-id="ef3b4-175">Test the add-in</span></span>

1. <span data-ttu-id="ef3b4-176">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-176">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="ef3b4-177">否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-177">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ef3b4-178">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-178">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="ef3b4-179">为此，需要终止服务器进程，这样才能看到提示并输入生成命令。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-179">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="ef3b4-180">生成后，重启服务器。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-180">After the build, restart the server.</span></span> <span data-ttu-id="ef3b4-181">接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-181">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="ef3b4-182">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-182">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="ef3b4-183">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-183">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="ef3b4-184">通过关闭任务窗格来重新加载它，再选择“开始”\*\*\*\* 菜单上的“显示任务窗格”\*\*\*\*，以重新打开加载项。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-184">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="ef3b4-185">在任务窗格中，选择“插入段落”\*\*\*\*，以确保文档开头有一个段落。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-185">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph at the start of the document.</span></span>
6. <span data-ttu-id="ef3b4-186">选择某文本。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-186">Select some text.</span></span> <span data-ttu-id="ef3b4-187">选择短语“即点即用”最合适。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-187">Selecting the phrase "Click-to-Run" will make the most sense.</span></span> <span data-ttu-id="ef3b4-188">*请注意，不要在选定区域的前后添加空格。*</span><span class="sxs-lookup"><span data-stu-id="ef3b4-188">*Be careful not to include the preceding or following space in the selection.*</span></span>
7. <span data-ttu-id="ef3b4-189">选择“插入缩写”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-189">Choose the **Insert Abbreviation** button.</span></span> <span data-ttu-id="ef3b4-190">观察“(C2R)”是否已添加。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-190">Note that " (C2R)" is added.</span></span> <span data-ttu-id="ef3b4-191">此外，还请观察，文档底部是否添加了包含整个扩展文本的新段落，因为新字符串已添加到现有区域中。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-191">Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span></span>
8. <span data-ttu-id="ef3b4-192">选择某文本。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-192">Select some text.</span></span> <span data-ttu-id="ef3b4-193">选择短语“Office 365”最合适。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-193">Selecting the phrase "Office 365" will make the most sense.</span></span> <span data-ttu-id="ef3b4-194">*请注意，不要在选定区域的前后添加空格。*</span><span class="sxs-lookup"><span data-stu-id="ef3b4-194">*Be careful not to include the preceding or following space in the selection.*</span></span>
9. <span data-ttu-id="ef3b4-195">选择“添加版本信息”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-195">Choose the **Add Version Info** button.</span></span> <span data-ttu-id="ef3b4-196">观察是否已在“Office 2016”和“Office 365”之间插入“Office 2019”。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-196">Note that "Office 2019, " is inserted between "Office 2016" and "Office 365".</span></span> <span data-ttu-id="ef3b4-197">此外，还请观察，文档底部是否添加了仅包含最初选定文本的新段落，因为新字符串已变成新区域，而不是添加到原始区域中。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-197">Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span></span>
10. <span data-ttu-id="ef3b4-198">选择某文本。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-198">Select some text.</span></span> <span data-ttu-id="ef3b4-199">选择字词“几个”最合适。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-199">Selecting the word "several" will make the most sense.</span></span> <span data-ttu-id="ef3b4-200">*请注意，不要在选定区域的前后添加空格。*</span><span class="sxs-lookup"><span data-stu-id="ef3b4-200">*Be careful not to include the preceding or following space in the selection.*</span></span>
11. <span data-ttu-id="ef3b4-201">选择“更改数量术语”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-201">Choose the **Change Quantity Term** button.</span></span> <span data-ttu-id="ef3b4-202">观察选定文本是否替换为“多个”。</span><span class="sxs-lookup"><span data-stu-id="ef3b4-202">Note that "many" replaces the selected text.</span></span>

    ![Word 教程 - 添加和替换文本](../images/word-tutorial-text-replace.png)
