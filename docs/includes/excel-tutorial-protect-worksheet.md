<span data-ttu-id="fa834-101">本教程的这一步是，向功能区添加另一个按钮。如果用户选择此按钮，便会执行所定义的函数，从而启用和禁用工作表保护。</span><span class="sxs-lookup"><span data-stu-id="fa834-101">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

> [!NOTE]
> <span data-ttu-id="fa834-102">此为 Excel 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="fa834-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="fa834-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="fa834-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="fa834-104">将清单配置为添加第二个功能区按钮</span><span class="sxs-lookup"><span data-stu-id="fa834-104">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="fa834-105">打开清单文件 **my-office-add-in-manifest.xml**。</span><span class="sxs-lookup"><span data-stu-id="fa834-105">Open the manifest file **my-office-add-in-manifest.xml**.</span></span>
2. <span data-ttu-id="fa834-106">找到 `<Control>` 元素。</span><span class="sxs-lookup"><span data-stu-id="fa834-106">Find the `<Control>` element.</span></span> <span data-ttu-id="fa834-107">此元素定义了“主页”\*\*\*\* 功能区上一直用于启动加载项的“显示任务窗格”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="fa834-107">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="fa834-108">将向“主页”\*\*\*\* 功能区上的相同组添加第二个按钮。</span><span class="sxs-lookup"><span data-stu-id="fa834-108">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="fa834-109">在结束 Control 标记 (`</Control>`) 和结束 Group 标记 (`</Group>`) 之间，添加下列标记。</span><span class="sxs-lookup"><span data-stu-id="fa834-109">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="fa834-110">将 `TODO1` 替换为字符串，以便向按钮提供在此清单文件内唯一的 ID。</span><span class="sxs-lookup"><span data-stu-id="fa834-110">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="fa834-111">因为清单中只有一个其他按钮，所以此操作并不难。</span><span class="sxs-lookup"><span data-stu-id="fa834-111">There's only one other button in the manifest, so this isn't difficult.</span></span> <span data-ttu-id="fa834-112">由于按钮将启用和禁用工作表保护，因此请使用“ToggleProtection”。</span><span class="sxs-lookup"><span data-stu-id="fa834-112">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="fa834-113">完成后，整个开始 Control 标记应如下所示：</span><span class="sxs-lookup"><span data-stu-id="fa834-113">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="fa834-114">接下来的三个 `TODO` 设置“resid”（这是资源 ID 的简称）。</span><span class="sxs-lookup"><span data-stu-id="fa834-114">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="fa834-115">资源是字符串，这三个字符串将在后续步骤中创建。</span><span class="sxs-lookup"><span data-stu-id="fa834-115">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="fa834-116">现在，需要向资源提供 ID。</span><span class="sxs-lookup"><span data-stu-id="fa834-116">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="fa834-117">虽然按钮标签应名为“切换保护”，但此字符串的 *ID* 应为“ProtectionButtonLabel”。因此，完成的 `Label` 元素应如下面的代码所示：</span><span class="sxs-lookup"><span data-stu-id="fa834-117">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="fa834-118">元素定义了按钮的工具提示。`SuperTip`</span><span class="sxs-lookup"><span data-stu-id="fa834-118">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="fa834-119">由于工具提示标题应与按钮标签相同，因此使用完全相同的资源 ID，即“ProtectionButtonLabel”。</span><span class="sxs-lookup"><span data-stu-id="fa834-119">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="fa834-120">工具提示说明为“单击即可启用和禁用工作表保护”。</span><span class="sxs-lookup"><span data-stu-id="fa834-120">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="fa834-121">不过，`ID` 应为“ProtectionButtonToolTip”。</span><span class="sxs-lookup"><span data-stu-id="fa834-121">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="fa834-122">因此，完成后，整个 `SuperTip` 标记应如下面的代码所示：</span><span class="sxs-lookup"><span data-stu-id="fa834-122">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="fa834-123">在生产加载项中，不建议对两个不同的按钮使用相同的图标；但为了简单起见，本教程将采用这样的做法。</span><span class="sxs-lookup"><span data-stu-id="fa834-123">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="fa834-124">因此，新 `Control` 中的 `Icon` 标记直接就是现有 `Control` 中 `Icon` 元素的副本。</span><span class="sxs-lookup"><span data-stu-id="fa834-124">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="fa834-125">虽然清单中现有原始 `Control` 元素内的 `Action` 元素的类型设置为 `ShowTaskpane`，但新按钮不会要打开任务窗格，而是要运行在后续步骤中创建的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="fa834-125">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="fa834-126">因此，将 `TODO5` 替换为 `ExecuteFunction`，即触发自定义函数的按钮的操作类型。</span><span class="sxs-lookup"><span data-stu-id="fa834-126">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="fa834-127">开始 `Action` 标记应如下面的代码所示：</span><span class="sxs-lookup"><span data-stu-id="fa834-127">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="fa834-128">原始 `Action` 元素的子元素指定任务窗格 ID，以及应当在任务窗格中打开的页面 URL。</span><span class="sxs-lookup"><span data-stu-id="fa834-128">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="fa834-129">不过，`ExecuteFunction` 类型的 `Action` 元素只有一个子元素，用于命名控件执行的函数。</span><span class="sxs-lookup"><span data-stu-id="fa834-129">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="fa834-130">此函数（名为 `toggleProtection`）将在后续步骤中创建。</span><span class="sxs-lookup"><span data-stu-id="fa834-130">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="fa834-131">因此，将 `TODO6` 替换为以下标记：</span><span class="sxs-lookup"><span data-stu-id="fa834-131">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="fa834-132">此时，整个 `Control` 标记应如下所示：</span><span class="sxs-lookup"><span data-stu-id="fa834-132">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="fa834-133">向下滚动到清单的 `Resources` 部分。</span><span class="sxs-lookup"><span data-stu-id="fa834-133">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="fa834-134">将下列标记添加为 `bt:ShortStrings` 元素的子级。</span><span class="sxs-lookup"><span data-stu-id="fa834-134">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="fa834-135">将下列标记添加为 `bt:LongStrings` 元素的子级。</span><span class="sxs-lookup"><span data-stu-id="fa834-135">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="fa834-136">请务必保存文件。</span><span class="sxs-lookup"><span data-stu-id="fa834-136">Be sure to save the file.</span></span>

## <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="fa834-137">创建工作表保护函数</span><span class="sxs-lookup"><span data-stu-id="fa834-137">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="fa834-138">打开文件 \function-file\function-file.js。</span><span class="sxs-lookup"><span data-stu-id="fa834-138">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="fa834-139">此文件已有立即调用函数表达式 (IIFE)。</span><span class="sxs-lookup"><span data-stu-id="fa834-139">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="fa834-140">由于不需要自定义初始化逻辑，因此分配到 `Office.initialize` 的函数的空主体保留不动。</span><span class="sxs-lookup"><span data-stu-id="fa834-140">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="fa834-141">（不过，请勿删除它。</span><span class="sxs-lookup"><span data-stu-id="fa834-141">(But do not delete it.</span></span> <span data-ttu-id="fa834-142">属性不得为空值或未定义。）*在 IIFE 之外*，添加下列代码。`Office.initialize`</span><span class="sxs-lookup"><span data-stu-id="fa834-142">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="fa834-143">请注意，我们向方法指定了 `args` 参数，因此方法的最后一行为 `args.completed`。</span><span class="sxs-lookup"><span data-stu-id="fa834-143">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="fa834-144">**ExecuteFunction** 类型的所有加载项命令都必须满足这项要求。</span><span class="sxs-lookup"><span data-stu-id="fa834-144">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="fa834-145">它会指示 Office 主机应用，函数已完成，且 UI 可以再次变成响应式。</span><span class="sxs-lookup"><span data-stu-id="fa834-145">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="fa834-146">将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="fa834-146">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="fa834-147">此代码使用处于标准切换模式的工作表对象 protection 属性。</span><span class="sxs-lookup"><span data-stu-id="fa834-147">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="fa834-148">将在下一部分中进行介绍。`TODO2`</span><span class="sxs-lookup"><span data-stu-id="fa834-148">The `TODO2` will be explained in the next section.</span></span>

    ```javascript
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="fa834-149">添加代码以将文档属性提取到任务窗格的脚本对象</span><span class="sxs-lookup"><span data-stu-id="fa834-149">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="fa834-150">在本系列教程前面的所有函数中，都是将命令排入队列，以对 Office 文档执行*写入*操作。</span><span class="sxs-lookup"><span data-stu-id="fa834-150">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="fa834-151">每个函数结束时都会调用 `context.sync()` 方法，从而将排入队列的命令发送到文档，以供执行。</span><span class="sxs-lookup"><span data-stu-id="fa834-151">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="fa834-152">不过，在上一步中添加的代码调用的是 `sheet.protection.protected` 属性，这与之前编写的函数明显不同，因为 `sheet` 对象只是任务窗格脚本中的代理对象。</span><span class="sxs-lookup"><span data-stu-id="fa834-152">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="fa834-153">它并不了解文档的实际保护状态，因此它的 `protection.protected` 属性无法有实值。</span><span class="sxs-lookup"><span data-stu-id="fa834-153">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="fa834-154">必须先从文档提取保护状态，再用它设置 `sheet.protection.protected` 值。</span><span class="sxs-lookup"><span data-stu-id="fa834-154">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="fa834-155">只有这样，才能调用 `sheet.protection.protected`，而不导致异常抛出。</span><span class="sxs-lookup"><span data-stu-id="fa834-155">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="fa834-156">此提取过程分为三步：</span><span class="sxs-lookup"><span data-stu-id="fa834-156">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="fa834-157">将命令排入队列，以加载（即提取）代码需要读取的属性。</span><span class="sxs-lookup"><span data-stu-id="fa834-157">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>
   2. <span data-ttu-id="fa834-158">调用上下文对象的 `sync`方法，从而向文档发送已排入队列的命令以供执行，并返回请求获取的信息。</span><span class="sxs-lookup"><span data-stu-id="fa834-158">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>
   3. <span data-ttu-id="fa834-159">由于 `sync` 是异步方法，因此请先确保它已完成，然后代码才能调用已提取的属性。</span><span class="sxs-lookup"><span data-stu-id="fa834-159">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="fa834-160">只要代码需要从 Office 文档*读取*信息，就必须完成这些步骤。</span><span class="sxs-lookup"><span data-stu-id="fa834-160">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="fa834-p112">在 `toggleProtection` 函数中，将 `TODO2` 替换为下列代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="fa834-p112">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="fa834-163">每个 Excel 对象都有 `load` 方法。</span><span class="sxs-lookup"><span data-stu-id="fa834-163">Every Excel object has a `load` method.</span></span> <span data-ttu-id="fa834-164">对于要在参数中读取的对象属性，将它们指定为逗号分隔名称字符串。</span><span class="sxs-lookup"><span data-stu-id="fa834-164">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="fa834-165">在此示例中，需要读取的属性为 `protection` 属性的子属性。</span><span class="sxs-lookup"><span data-stu-id="fa834-165">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="fa834-166">引用子属性的方法与在代码中的其他任何地方引用属性几乎完全一样，不同之处在于使用的是正斜杠（“/”）字符，而不是“.”字符。</span><span class="sxs-lookup"><span data-stu-id="fa834-166">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>
   - <span data-ttu-id="fa834-167">为了确保切换逻辑 `sheet.protection.protected` 只在 `sync` 完成后且 `sheet.protection.protected` 分配有从文档提取的正确值后才运行，（在下一步中）它会被移到 `then` 函数中，此函数在 `sync` 完成前不会运行。</span><span class="sxs-lookup"><span data-stu-id="fa834-167">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```javascript
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="fa834-168">由于不能在同一取消分支代码路径中有两个 `return` 语句，因此请删除 `Excel.run` 末尾的最后一行代码 `return context.sync();`。</span><span class="sxs-lookup"><span data-stu-id="fa834-168">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="fa834-169">新的最后一行代码 `context.sync`将在后续步骤中添加。</span><span class="sxs-lookup"><span data-stu-id="fa834-169">You will add a new final `context.sync`, in a later step.</span></span>
3. <span data-ttu-id="fa834-170">剪切并粘贴 `toggleProtection` 函数中的 `if ... else` 结构，以替换 `TODO3`。</span><span class="sxs-lookup"><span data-stu-id="fa834-170">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>
4. <span data-ttu-id="fa834-p115">将 `TODO4` 替换为以下代码。注意：</span><span class="sxs-lookup"><span data-stu-id="fa834-p115">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="fa834-173">将 `sync` 方法传递到 `then` 函数可确保它不会在 `sheet.protection.unprotect()` 或 `sheet.protection.protect()` 已排入队列前运行。</span><span class="sxs-lookup"><span data-stu-id="fa834-173">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>
   - <span data-ttu-id="fa834-174">由于 `then` 方法调用传递给它的任何函数，并且也不想调用 `sync` 两次，因此请从 `context.sync` 末尾省略掉“()”。</span><span class="sxs-lookup"><span data-stu-id="fa834-174">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```javascript
    .then(context.sync);
    ```

   <span data-ttu-id="fa834-175">完成后，整个函数应如下所示：</span><span class="sxs-lookup"><span data-stu-id="fa834-175">When you are done, the entire function should look like the following:</span></span>

    ```javascript
    function toggleProtection(args) {
        Excel.run(function (context) {            
          const sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
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
        args.completed();
    }
    ```


## <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="fa834-176">配置脚本加载 HTML 文件</span><span class="sxs-lookup"><span data-stu-id="fa834-176">Configure the script-loading HTML file</span></span>

<span data-ttu-id="fa834-177">打开 /function-file/function-file.html 文件。</span><span class="sxs-lookup"><span data-stu-id="fa834-177">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="fa834-178">这是在用户按“切换工作表保护”\*\*\*\* 按钮时调用的无 UI HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="fa834-178">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="fa834-179">用于加载应当在按钮按下时运行的 JavaScript 方法。</span><span class="sxs-lookup"><span data-stu-id="fa834-179">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="fa834-180">将不更改此文件。</span><span class="sxs-lookup"><span data-stu-id="fa834-180">You are not going to change this file.</span></span> <span data-ttu-id="fa834-181">只需注意，第二个 `<script>` 标记加载 functionfile.js。</span><span class="sxs-lookup"><span data-stu-id="fa834-181">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="fa834-182">function-file.html 文件及其加载的 function-file.js 文件在完全独立于加载项任务窗格的 IE 进程中运行。</span><span class="sxs-lookup"><span data-stu-id="fa834-182">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="fa834-183">如果将 function-file.js 转换为与 app.js 文件相同的 bundle.js 文件，加载项必须加载 bundle.js 文件的两个副本，这就违背了绑定目的。</span><span class="sxs-lookup"><span data-stu-id="fa834-183">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="fa834-184">此外，function-file.js 文件不包含任何不受 IE 支持的 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="fa834-184">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="fa834-185">出于这两点原因，此加载项根本不会转换 function-file.js。</span><span class="sxs-lookup"><span data-stu-id="fa834-185">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="fa834-186">测试加载项</span><span class="sxs-lookup"><span data-stu-id="fa834-186">Test the add-in</span></span>

1. <span data-ttu-id="fa834-187">关闭包括 Excel 在内的所有 Office 应用。</span><span class="sxs-lookup"><span data-stu-id="fa834-187">Close all Office applications, including Excel.</span></span> 
2. <span data-ttu-id="fa834-188">通过删除缓存文件夹内容，删除 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="fa834-188">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="fa834-189">若要完全清除主机中的旧版加载项，必须这样做。</span><span class="sxs-lookup"><span data-stu-id="fa834-189">This is necessary to completely clear the old version of the add-in from the host.</span></span> 
    - <span data-ttu-id="fa834-190">对于 Windows：`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。</span><span class="sxs-lookup"><span data-stu-id="fa834-190">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>
    - <span data-ttu-id="fa834-191">对于 Mac：`/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`。</span><span class="sxs-lookup"><span data-stu-id="fa834-191">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>
3. <span data-ttu-id="fa834-192">如果服务器出于任何原因而未运行，请在 Git Bash 窗口或已启用 Node.JS 的系统命令提示符中，转到项目的“开始”\*\*\*\* 文件夹，再运行命令 `npm start`。</span><span class="sxs-lookup"><span data-stu-id="fa834-192">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="fa834-193">无需重新生成项目，因为唯一更改的 JavaScript 文件不属于已生成的 bundle.js。</span><span class="sxs-lookup"><span data-stu-id="fa834-193">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>
4. <span data-ttu-id="fa834-194">使用更改后的新版清单文件，并通过下列方法之一，重复旁加载进程。</span><span class="sxs-lookup"><span data-stu-id="fa834-194">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="fa834-195">*应覆盖清单文件的旧副本。*</span><span class="sxs-lookup"><span data-stu-id="fa834-195">*You should overwrite the previous copy of the manifest file.*</span></span>
    - <span data-ttu-id="fa834-196">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="fa834-196">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="fa834-197">Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="fa834-197">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="fa834-198">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="fa834-198">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
7. <span data-ttu-id="fa834-199">打开 Excel 中的任意工作表。</span><span class="sxs-lookup"><span data-stu-id="fa834-199">Open any worksheet in Excel.</span></span>
8. <span data-ttu-id="fa834-p121">在“开始”\*\*\*\* 功能区上，选择“切换工作表保护”\*\*\*\*。请注意，功能区上的大部分控件都处于禁用状态（灰显），如下面的屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="fa834-p121">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 
9. <span data-ttu-id="fa834-202">选择要更改其内容的单元格。</span><span class="sxs-lookup"><span data-stu-id="fa834-202">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="fa834-203">此时，将会看到一条错误消息，提示工作表受保护。</span><span class="sxs-lookup"><span data-stu-id="fa834-203">You get an error telling you that the worksheet is protected.</span></span>
10. <span data-ttu-id="fa834-204">再次选择“切换工作表保护”\*\*\*\*，此时控件重新启用，可以再次更改单元格值了。</span><span class="sxs-lookup"><span data-stu-id="fa834-204">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Excel 教程 - 在功能区上启用工作表保护](../images/excel-tutorial-ribbon-with-protection-on.png)
