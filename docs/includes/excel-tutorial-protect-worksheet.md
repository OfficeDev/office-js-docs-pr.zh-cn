本教程的这一步是，向功能区添加另一个按钮。如果用户选择此按钮，便会执行所定义的函数，从而启用和禁用工作表保护。

> [!NOTE]
> 此为 Excel 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a>将清单配置为添加第二个功能区按钮

1. 打开清单文件 **my-office-add-in-manifest.xml**。
2. 找到 `<Control>` 元素。 此元素定义了“主页”功能区上一直用于启动加载项的“显示任务窗格”按钮。 将向“主页”功能区上的相同组添加第二个按钮。 在结束 Control 标记 (`</Control>`) 和结束 Group 标记 (`</Group>`) 之间，添加下列标记。

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

3. 将 `TODO1` 替换为字符串，以便向按钮提供在此清单文件内唯一的 ID。 因为清单中只有一个其他按钮，所以此操作并不难。 由于按钮将启用和禁用工作表保护，因此请使用“ToggleProtection”。 完成后，整个开始 Control 标记应如下所示：

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. 接下来的三个 `TODO` 设置“resid”（这是资源 ID 的简称）。 资源是字符串，这三个字符串将在后续步骤中创建。 现在，需要向资源提供 ID。 虽然按钮标签应名为“切换保护”，但此字符串的 *ID* 应为“ProtectionButtonLabel”。因此，完成的 `Label` 元素应如下面的代码所示：

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. `SuperTip` 元素定义了按钮的工具提示。 由于工具提示标题应与按钮标签相同，因此使用完全相同的资源 ID，即“ProtectionButtonLabel”。 工具提示说明为“单击即可启用和禁用工作表保护”。 不过，`ID` 应为“ProtectionButtonToolTip”。 因此，完成后，整个 `SuperTip` 标记应如下面的代码所示： 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > 在生产加载项中，不建议对两个不同的按钮使用相同的图标；但为了简单起见，本教程将采用这样的做法。 因此，新 `Control` 中的 `Icon` 标记直接就是现有 `Control` 中 `Icon` 元素的副本。 

6. 虽然清单中现有原始 `Control` 元素内的 `Action` 元素的类型设置为 `ShowTaskpane`，但新按钮不会要打开任务窗格，而是要运行在后续步骤中创建的自定义函数。 因此，将 `TODO5` 替换为 `ExecuteFunction`，即触发自定义函数的按钮的操作类型。 开始 `Action` 标记应如下面的代码所示：
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. 原始 `Action` 元素的子元素指定任务窗格 ID，以及应当在任务窗格中打开的页面 URL。 不过，`ExecuteFunction` 类型的 `Action` 元素只有一个子元素，用于命名控件执行的函数。 此函数（名为 `toggleProtection`）将在后续步骤中创建。 因此，将 `TODO6` 替换为以下标记：
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    此时，整个 `Control` 标记应如下所示：

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

8. 向下滚动到清单的 `Resources` 部分。

9. 将下列标记添加为 `bt:ShortStrings` 元素的子级。

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. 将下列标记添加为 `bt:LongStrings` 元素的子级。

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. 请务必保存文件。

## <a name="create-the-function-that-protects-the-sheet"></a>创建工作表保护函数

1. 打开文件 \function-file\function-file.js。

2. 此文件已有立即调用函数表达式 (IIFE)。 由于不需要自定义初始化逻辑，因此分配到 `Office.initialize` 的函数的空主体保留不动。 （不过，请勿删除它。 `Office.initialize` 属性不得为空值或未定义。）*在 IIFE 之外*，添加下列代码。 请注意，我们向方法指定了 `args` 参数，因此方法的最后一行为 `args.completed`。 **ExecuteFunction** 类型的所有加载项命令都必须满足这项要求。 它会指示 Office 主机应用，函数已完成，且 UI 可以再次变成响应式。

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

3. 将 `TODO1` 替换为以下代码。 此代码使用处于标准切换模式的工作表对象 protection 属性。 `TODO2` 将在下一部分中进行介绍。

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

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>添加代码以将文档属性提取到任务窗格的脚本对象

在本系列教程前面的所有函数中，都是将命令排入队列，以对 Office 文档执行*写入*操作。 每个函数结束时都会调用 `context.sync()` 方法，从而将排入队列的命令发送到文档，以供执行。 不过，在上一步中添加的代码调用的是 `sheet.protection.protected` 属性，这与之前编写的函数明显不同，因为 `sheet` 对象只是任务窗格脚本中的代理对象。 它并不了解文档的实际保护状态，因此它的 `protection.protected` 属性无法有实值。 必须先从文档提取保护状态，再用它设置 `sheet.protection.protected` 值。 只有这样，才能调用 `sheet.protection.protected`，而不导致异常抛出。 此提取过程分为三步：

   1. 将命令排入队列，以加载（即提取）代码需要读取的属性。
   2. 调用上下文对象的 `sync`方法，从而向文档发送已排入队列的命令以供执行，并返回请求获取的信息。
   3. 由于 `sync` 是异步方法，因此请先确保它已完成，然后代码才能调用已提取的属性。

只要代码需要从 Office 文档*读取*信息，就必须完成这些步骤。

1. 在 `toggleProtection` 函数中，将 `TODO2` 替换为下列代码。请注意以下几点：
   - 每个 Excel 对象都有 `load` 方法。 对于要在参数中读取的对象属性，将它们指定为逗号分隔名称字符串。 在此示例中，需要读取的属性为 `protection` 属性的子属性。 引用子属性的方法与在代码中的其他任何地方引用属性几乎完全一样，不同之处在于使用的是正斜杠（“/”）字符，而不是“.”字符。
   - 为了确保切换逻辑 `sheet.protection.protected` 只在 `sync` 完成后且 `sheet.protection.protected` 分配有从文档提取的正确值后才运行，（在下一步中）它会被移到 `then` 函数中，此函数在 `sync` 完成前不会运行。 

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

2. 由于不能在同一取消分支代码路径中有两个 `return` 语句，因此请删除 `Excel.run` 末尾的最后一行代码 `return context.sync();`。 新的最后一行代码 `context.sync`将在后续步骤中添加。
3. 剪切并粘贴 `toggleProtection` 函数中的 `if ... else` 结构，以替换 `TODO3`。
4. 将 `TODO4` 替换为以下代码。注意：
   - 将 `sync` 方法传递到 `then` 函数可确保它不会在 `sheet.protection.unprotect()` 或 `sheet.protection.protect()` 已排入队列前运行。
   - 由于 `then` 方法调用传递给它的任何函数，并且也不想调用 `sync` 两次，因此请从 `context.sync` 末尾省略掉“()”。

    ```javascript
    .then(context.sync);
    ```

   完成后，整个函数应如下所示：

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


## <a name="configure-the-script-loading-html-file"></a>配置脚本加载 HTML 文件

打开 /function-file/function-file.html 文件。 这是在用户按“切换工作表保护”按钮时调用的无 UI HTML 文件。 用于加载应当在按钮按下时运行的 JavaScript 方法。 将不更改此文件。 只需注意，第二个 `<script>` 标记加载 functionfile.js。

   > [!NOTE]
   > function-file.html 文件及其加载的 function-file.js 文件在完全独立于加载项任务窗格的 IE 进程中运行。 如果将 function-file.js 转换为与 app.js 文件相同的 bundle.js 文件，加载项必须加载 bundle.js 文件的两个副本，这就违背了绑定目的。 此外，function-file.js 文件不包含任何不受 IE 支持的 JavaScript。 出于这两点原因，此加载项根本不会转换 function-file.js。 

## <a name="test-the-add-in"></a>测试加载项

1. 关闭包括 Excel 在内的所有 Office 应用。 
2. 通过删除缓存文件夹内容，删除 Office 缓存。 若要完全清除主机中的旧版加载项，必须这样做。 
    - 对于 Windows：`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。
    - 对于 Mac：`/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`。
3. 如果服务器出于任何原因而未运行，请在 Git Bash 窗口或已启用 Node.JS 的系统命令提示符中，转到项目的“开始”文件夹，再运行命令 `npm start`。 无需重新生成项目，因为唯一更改的 JavaScript 文件不属于已生成的 bundle.js。
4. 使用更改后的新版清单文件，并通过下列方法之一，重复旁加载进程。 *应覆盖清单文件的旧副本。*
    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
7. 打开 Excel 中的任意工作表。
8. 在“开始”功能区上，选择“切换工作表保护”。请注意，功能区上的大部分控件都处于禁用状态（灰显），如下面的屏幕截图所示。 
9. 选择要更改其内容的单元格。 此时，将会看到一条错误消息，提示工作表受保护。
10. 再次选择“切换工作表保护”，此时控件重新启用，可以再次更改单元格值了。

    ![Excel 教程 - 在功能区上启用工作表保护](../images/excel-tutorial-ribbon-with-protection-on.png)
