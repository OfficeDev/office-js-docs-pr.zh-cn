---
title: Word 加载项教程
description: 本教程将介绍如何生成 Word 加载项，用于插入（和替换）文本区域、段落、图像、HTML、表格和内容控件。 此外，还将介绍如何设置文本格式，以及如何插入（和替换）内容控件中的内容。
ms.date: 01/13/2022
ms.prod: word
ms.localizationpriority: high
ms.openlocfilehash: 1f7950007a9139767cd31901ccf64c9fb1ebdf7c
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958381"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a>教程：创建 Word 任务窗格加载项

在本教程中，将创建 Word 任务窗格加载项，该加载项将：

> [!div class="checklist"]
>
> - 插入文本区域
> - 设置文本格式
> - 替换文本并在各个位置插入文本
> - 插入图像、HTML 和表格
> - 创建和更新内容控件

> [!TIP]
> 如果已完成了[创建首个 Word 任务窗格加载项](../quickstarts/word-quickstart.md)快速入门，并希望使用该项目作为本教程的起点，请直接转到[插入文本区域](#insert-a-range-of-text)以开始此教程。

## <a name="prerequisites"></a>先决条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- 已连接到 Microsoft 365 订阅的 Office (包括 Office 网页版)。

    > [!NOTE]
    > 如果你还没有 Office，可以[加入 Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)以免费获得为期 90 天的可续订 Microsoft 365 订阅，以便在开发期间使用。

## <a name="create-your-add-in-project"></a>创建加载项项目

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **选择项目类型:** `Office Add-in Task Pane project`
- **选择脚本类型:** `Javascript`
- **要如何命名加载项?** `My Office Add-in`
- **要支持哪一个 Office 客户端应用程序?** `Word`

![显示命令行界面中 Yeoman 生成器的提示和回答的屏幕截图。](../images/yo-office-word.png)

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="insert-a-range-of-text"></a>插入文本区域

本教程的这一步是，先以编程方式测试加载项是否支持用户的当前版本 Word，再在文档中插入段落。

### <a name="code-the-add-in"></a>编码加载项

1. 在代码编辑器中打开项目。

1. 打开文件 **./src/taskpane/taskpane.html**。此文件包含任务窗格的 HTML 标记。

1. 找到 `<main>` 元素并删除在开始 `<main>` 标记后和关闭 `</main>` 标记前出现的所有行。

1. 打开 `<main>` 标记后立即添加下列标记：

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button><br/><br/>
    ```

1. 打开文件 **./src/taskpane/taskpane.js**。此文件包含可促进任务窗格和 Office 客户端应用程序之间交互的 Office JavaScript API 代码。

1. 执行以下操作，删除对 `run` 按钮和 `run()` 函数的所有引用：

    - 查找并删除行 `document.getElementById("run").onclick = run;`。

    - 查找并删除整个 `run()` 函数。

1. 在 `Office.onReady` 函数调用中，找到行 `if (info.host === Office.HostType.Word) {` 并紧跟该行添加下列代码。 注意：

    - 此代码的第一部分用于确定用户的 Word 版本是否支持包含本教程所有阶段使用的全部 API 的 Word.js 版本。在生产加载项中，若要隐藏或禁用调用不受支持的 API 的 UI，请使用条件块的主体。这样一来，用户仍可以使用 Word 版本支持的加载项部分。
    - 此代码的第二部分为 `insert-paragraph` 按钮添加了事件处理程序。

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    ```

1. 将以下函数添加到文件结尾。注意：

   - Word.js 业务逻辑会添加到传递给 `Word.run` 的函数中。 此逻辑不会立即执行， 而是添加到挂起命令队列中。

   - `context.sync` 方法将所有已排入队列的命令都发送到 Word 以供执行。

   - `Word.run` 后跟 `catch` 块。 这是应始终遵循的最佳做法。

   [!include[Information about the use of ES6 JavaScript](../includes/modern-js-note.md)]

    ```js
    async function insertParagraph() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a paragraph into the document.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `insertParagraph()` 函数中，将 `TODO1` 替换为下列代码。注意：

   - `insertParagraph` 方法的第一个参数是新段落的文本。

   - 第二个参数是应在正文中的什么位置插入段落。 如果父对象为正文，其他段落插入选项包括“End”和“Replace”。

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
                            "Start");
    ```

1. 验证是否已保存了对项目所做的所有更改。

### <a name="test-the-add-in"></a>测试加载项

1. 完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > 如果在 Mac 上测试加载项，请先运行项目根目录中的以下命令，然后再继续。 运行此命令时，本地 Web 服务器将启动。
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - 若要在 Word 中测试加载项，请在项目的根目录中运行以下命令。 这将启动本地的 Web 服务器（如果尚未运行的话），并使用加载的加载项打开 Word。

        ```command&nbsp;line
        npm start
        ```

    - 若要在 Word 网页版中测试加载项，请在项目的根目录中运行以下命令。 运行此命令时，本地 Web 服务器将启动。 将“{url}”替换为你有权访问的 OneDrive 或 SharePoint 库中 Word 文档的 URL。

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. 在 Word 中，依次选择“开始”选项卡和功能区中的“显示任务窗格”按钮，以打开加载项任务窗格。

    ![显示 Word 中突出显示的“显示任务窗格”按钮的屏幕截图。](../images/word-quickstart-addin-2b.png)

1. 在任务窗格中，选择“插入段落”按钮。

1. 在段落中进行一些更改。

1. 再次选择“**插入段落**”按钮。请注意，新段落出现在上一个段落的上方，因为 `insertParagraph` 方法将在文档正文的开头插入。

    ![显示加载项中“插入段落”按钮的屏幕截图。](../images/word-tutorial-insert-paragraph-2.png)

## <a name="format-text"></a>设置文本格式

在本教程的此步骤中，你将向文本应用嵌入样式、向文本应用自定义样式并更改文本字体。

### <a name="apply-a-built-in-style-to-text"></a>向文本应用嵌入样式

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`insert-paragraph`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="apply-style">Apply Style</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `insert-paragraph` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("apply-style").onclick = applyStyle;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function applyStyle() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to style text.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `applyStyle()` 函数中，将 `TODO1` 替换为以下代码。请注意，代码将样式应用于段落，但样式也可以应用于文本范围。

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ```

### <a name="apply-a-custom-style-to-text"></a>向文本应用自定义样式

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`apply-style`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `apply-style` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function applyCustomStyle() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to apply the custom style.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在`applyCustomStyle()`函数中，用`TODO1`替换代码代码。请注意，代码应用了一个尚不存在的自定义样式。你将在 [测试外接程序](#test-the-add-in-1)步骤中创建名称为 **MyCustomStyle** 的样式。

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ```

1. 验证是否已保存了对项目所做的所有更改。

### <a name="change-the-font-of-text"></a>更改文本字体

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`apply-custom-style`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="change-font">Change Font</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `apply-custom-style` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("change-font").onclick = changeFont;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function changeFont() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to apply a different font.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在`changeFont()`函数中，将`TODO1`替换为以下代码。请注意，代码通过使用链接到`Paragraph.getNext`方法的`ParagraphCollection.getFirst`方法获取对第二个段落的引用。

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ```

1. 验证是否已保存了对项目所做的所有更改。

### <a name="test-the-add-in"></a>测试加载项

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. 如果加载项任务窗格已在 Word 中打开，请转到“开始”选项卡并选择功能区中的“显示任务窗格”按钮以打开它。

1. 请确保文档中至少有三个段落。 可以选择“插入段落”按钮三次。 仔细检查文档末尾是否没有空白段落。若有，请删除它。

1. 在 Word 中，创建[自定义样式](https://support.microsoft.com/office/d38d6e47-f6fc-48eb-a607-1eb120dec563)“MyCustomStyle”。其中可以包含所需的任何格式。

1. 选择 **“应用样式”** 按钮。 第一个段落将采用嵌入样式 **“明显参考”**。

1. 选择 **“应用自定义样式”** 按钮。 最后一个段落将采用自定义样式。 （如果好像什么都没有发生，很可能是因为最后一个段落是空白段落。 如果是这样，请向其中添加某文本。）

1. 选择 **“更改字体”** 按钮。 第二个段落的字体更改为 18 磅的粗体 Courier New。

    ![显示为加载项按钮“应用样式”、“应用自定义样式”和“更改字体”应用了定义的样式和字体的结果屏幕截图。](../images/word-tutorial-apply-styles-and-font-2.png)

## <a name="replace-text-and-insert-text"></a>替换文本和插入文本

本教程的这一步是，在选定文本区域内外添加文本，并替换选定区域的文本。

### <a name="add-text-inside-a-range"></a>在区域内添加文本

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`change-font`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `change-font` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function insertTextIntoRange() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `insertTextIntoRange()` 函数中，将 `TODO1` 替换为以下代码。注意：

   - 此函数用于在“即点即用”文本区域末尾插入缩写 [“(C2R)”]。 它做了一个简化假设，即存在字符串，且用户已选择它。

   - `Range.insertText` 方法的第一个参数是要插入到 `Range` 对象的字符串。

   - 第二个参数指定了应在区域中的什么位置插入其他文本。 除了“End”外，其他可用选项包括“Start”、“Before”、“After”和“Replace”。

   - “End”和“After”的区别在于，“End”在现有区域末尾插入新文本，而“After”则是新建包含字符串的区域，并在现有区域后面插入新区域。 同样，“Start”是在现有区域的开头位置插入文本，而“Before”插入的是新区域。 “Replace”将现有区域文本替换为第一个参数中的字符串。

   - 在本教程之前阶段中，正文对象的 insert* 方法没有“Before”和“After”选项。 这是因为不能将内容置于文档正文外。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

1. 我们将跳过 `TODO2`，直接到下一部分。在 `insertTextIntoRange()` 函数中，将 `TODO3` 替换为以下代码。此代码类似于在本教程的第一阶段中创建的代码，只是现在在文档末尾（而不是在开头）插入新段落。此新段落将演示新文本现在是原始范围的一部分。

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>添加代码以将文档属性提取到任务窗格的脚本对象

在本系列教程前面的所有函数中，都是将命令排入队列，以 *将* 写入 Office 文档。每个函数结束时都会调用 `context.sync()` 方法，从而将排入队列的命令发送到文档，以供执行。不过，在上一步中添加的代码调用的是 `originalRange.text` 属性，这与之前编写的函数明显不同，因为 `originalRange` 对象只是任务窗格脚本中的代理对象。由于它并不了解文档中区域的实际文本，因此它的 `text` 属性无法有实际值。有必要先从文档中提取区域的文本值，再用它设置  `originalRange.text` 的值。 只有这样才能调用 `originalRange.text`，而又不会引发异常。 此提取过程分为三步：

1. 将命令排入队列，以加载 (即提取) 代码需要读取的属性。

1. 调用上下文对象的 `sync`方法，从而向文档发送已排入队列的命令以供执行，并返回请求获取的信息。

1. 由于 `sync` 是异步方法，因此请先确保它已完成，然后代码才能调用已提取的属性。

只要代码需要从 Office 文档 *读取* 信息，就必须完成这些步骤。

1. 在 `insertTextIntoRange()` 函数中，将 `TODO2` 替换为以下代码。
  
    ```js
    originalRange.load("text");
    await context.sync();

    // TODO4: Move the doc.body.insertParagraph line here.

    // TODO5: Move the final call of context.sync here and ensure
    //        that it does not run until the insertParagraph has
    //        been queued.
    ```

1. 剪切并粘贴 `doc.body.insertParagraph` 代码行，以替代 `TODO4`。

完成后，整个函数应如下所示：

```js
async function insertTextIntoRange() {
    await Word.run(async (context) => {

        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        await context.sync();

        doc.body.insertParagraph("Original range: " + originalRange.text, "End");

        await context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
```

### <a name="add-text-between-ranges"></a>在区域间添加文本

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`insert-text-into-range`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `insert-text-into-range` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function insertTextBeforeRange() {
        await Word.run(async (context) => {

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

1. 在 `insertTextBeforeRange()` 函数中，将 `TODO1` 替换为以下代码。注意：

   - 此函数用于带有文本“Microsoft 365”的区域前添加文本为“Office 2019”的区域。 它做了一个简化假设，即存在字符串，且用户已选择它。

   - `Range.insertText` 方法的第一个参数是要添加的字符串。

   - 第二个参数指定了应在区域中的什么位置插入其他文本。 若要详细了解位置选项，请参阅前面介绍的 `insertTextIntoRange` 函数。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

1. 在 `insertTextBeforeRange()` 函数中，将 `TODO2` 替换为以下代码。

     ```js
    originalRange.load("text");
    await context.sync();

    // TODO3: Queue commands to insert the original range as a
    //        paragraph at the end of the document.

    // TODO4: Make a final call of context.sync here and ensure
    //        that it runs after the insertParagraph has been queued.
    ```

1. 将 `TODO3` 替换为下面的代码。 这一新段落将说明，新文本 ***不*** 属于原始选定区域。 原始区域中的文本仍与用户选择它时一样。

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
    ```

1. 将 `TODO4` 替换为下面的代码。

    ```js
    await context.sync();
    ```

### <a name="replace-the-text-of-a-range"></a>替换区域文本

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`insert-text-outside-range`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="replace-text">Change Quantity Term</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `insert-text-outside-range` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("replace-text").onclick = replaceText;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function replaceText() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to replace the text.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `replaceText()` 函数中，将 `TODO1` 替换为以下代码。 请注意，此函数用于将字符串“几个”替换为字符串“许多”。 它做了一个简化假设，即存在字符串，且用户已选择它。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

1. 验证是否已保存了对项目所做的所有更改。

### <a name="test-the-add-in"></a>测试加载项

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. 如果加载项任务窗格已在 Word 中打开，请转到“开始”选项卡并选择功能区中的“显示任务窗格”按钮以打开它。

1. 在任务窗格中，选择“插入段落”按钮，以确保文档开头有一个段落。

1. 在文档中，选择短语“即点即用”。*注意在选择时不要包括前面的空格或后面的逗号。*

1. 选择 **“插入缩写”** 按钮。 观察“(C2R)”是否已添加。 此外，还请观察，文档底部是否添加了包含整个扩展文本的新段落，因为新字符串已添加到现有区域中。

1. 在文档中，选择短语“Microsoft 365”。*注意不要在所选内容中包含前导或尾随空格。*

1. 选择“**添加版本信息**” 按钮。观察是否已在“Office 2016”和“Microsoft 365”之间插入“Office 2019”。此外，还请观察，文档底部是否添加了仅包含最初选定文本的新段落，因为新字符串已变成新区域，而不是添加到原始区域中。

1. 在文档中，选择单词“several”。 *请注意不要在所选内容中包含前面或以下空格。*

1. 选择 **“更改数量术语”** 按钮。观察选定文本是否替换为“多个”。

    ![屏幕截图显示选择加载项按钮的结果“插入缩写”、“添加版本信息”和“更改数量术语”。](../images/word-tutorial-text-replace-2.png)

## <a name="insert-images-html-and-tables"></a>插入图像、HTML 和表格

本教程的这一步是，了解如何在文档中插入图像、HTML 和表格。

### <a name="define-an-image"></a>定义图像

完成以下步骤，定义要在本教程的下一部分插入到文档中的图像。

1. 在项目的根目录中，创建一个名为 base64Image.js 的新文件。

1. 打开文件 base64Image.js 并添加以下代码，以指定表示图像的 base64 编码字符串。

    ```js
    export const base64Image =
        "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";
    ```

### <a name="insert-an-image"></a>插入图像

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`replace-text`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="insert-image">Insert Image</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在文件顶部附近找到 `Office.onReady` 函数调用，然后在该行之前添加以下代码。 此代码将导入你先前在文件 /base64Image.js 中定义的变量。

    ```js
    import { base64Image } from "../../base64Image";
    ```

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `replace-text` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("insert-image").onclick = insertImage;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function insertImage() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert an image.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `insertImage()` 函数中，将 `TODO1` 替换为以下代码。请注意，此代码行在文档末尾插入 Base64 编码图像。（`Paragraph` 对象还包含 `insertInlinePictureFromBase64` 方法和其他 `insert*` 方法。有关示例，请参阅以下 insertHTML 部分。）

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a>插入 HTML

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`insert-image`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="insert-html">Insert HTML</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `insert-image` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("insert-html").onclick = insertHTML;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function insertHTML() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to insert a string of HTML.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `insertHTML()` 函数中，将 `TODO1` 替换为以下代码。注意：

   - 第一行代码在文档末尾添加空白段落。

   - 第二行代码在段落末尾插入 HTML 字符串；具体而言是两个段落，一个设置使用 Verdana 字体格式，另一个采用 Word 文档的默认样式。 （如前面的 `insertImage` 方法一样，`context.document.body` 对象还包含 `insert*` 方法。）

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a>插入表格

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`insert-html`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="insert-table">Insert Table</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `insert-html` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("insert-table").onclick = insertTable;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function insertTable() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `insertTable()` 函数中，将 `TODO1` 替换为以下代码。请注意，此行使用 `ParagraphCollection.getFirst` 方法获取对第一个段落的引用，然后使用 `Paragraph.getNext` 方法获取对第二个段落的引用。

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

1. 在 `insertTable()` 函数中，将 `TODO2` 替换为以下代码。注意：

   - `insertTable` 方法的前两个参数指定行数和列数。

   - 第三个参数指定要在哪里插入表格（在此示例中，是在段落后面插入）。

   - 第四个参数是用于设置表格单元格值的二维数组。

   - 虽然表格采用普通的默认样式，但 `insertTable` 方法返回的 `Table` 对象包含多个成员，其中部分成员用于设置表格样式。

    ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

1. 验证是否已保存了对项目所做的所有更改。

### <a name="test-the-add-in"></a>测试加载项

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. 如果加载项任务窗格已在 Word 中打开，请转到“开始”选项卡并选择功能区中的“显示任务窗格”按钮以打开它。

1. 在任务窗格中，至少选择“插入段落”按钮三次，以确保文档中有多个段落。

1. 选择“插入图像”按钮，观察图像是否插入在文档末尾。

1. 选择“插入 HTML”按钮，观察是否在文档末尾插入了两个段落，第一个段落使用 Verdana 字体。

1. 选择“插入表格”按钮，观察是否在第二个段落后面插入了表格。

    ![显示选择加载项按钮“插入图像”、“插入 HTML”和“插入表”结果的屏幕截图。](../images/word-tutorial-insert-image-html-table-2.png)

## <a name="create-and-update-content-controls"></a>创建和更新内容控件

本教程的这一步是，了解如何在文档中创建格式文本内容控件，以及如何插入和替换控件的内容。

> [!NOTE]
> 虽然可通过 UI 添加到 Word 文档的内容控件有好几种，但目前 Word.js 仅支持格式文本内容控件。
>
> 开始执行本教程的这一步之前，建议通过 Word UI 创建和控制格式文本内容控件，以便熟悉此类控件及其属性。 有关详细信息，请参阅[在 Word 中创建用户填写或打印的表单](https://support.microsoft.com/office/040c5cc1-e309-445b-94ac-542f732c8c8b)。

### <a name="create-a-content-control"></a>创建内容控件

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`insert-table`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="create-content-control">Create Content Control</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `insert-table` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("create-content-control").onclick = createContentControl;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function createContentControl() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to create a content control.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `createContentControl()` 函数中，将 `TODO1` 替换为以下代码。注意：

   - 此代码旨在将短语“Microsoft 365”包装到内容控件中。它做了一个简化假设，即存在字符串，且用户已选择它。

   - `ContentControl.title` 属性指定内容控件的可见标题。

   - `ContentControl.tag` 属性指定标记，可用于通过 `ContentControlCollection.getByTag` 方法获取对内容控件的引用，将用于稍后出现的函数。

   - `ContentControl.appearance` 属性指定控件的外观。 使用值“Tags”表示，控件包装在开始标记和结束标记中，且开始标记包含内容控件标题。 其他可取值包括“BoundingBox”和“None”。

   - `ContentControl.color` 属性指定标记颜色或边界框的边框。

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a>替换内容控件的内容

1. 打开 ./src/taskpane/taskpane.html 文件。

1. 查找`create-content-control`按钮的`<button>`元素，并在行后添加下列标记。

    ```html
    <button class="ms-Button" id="replace-content-in-control">Rename Service</button><br/><br/>
    ```

1. 打开 **./src/taskpane/taskpane.js** 文件。

1. 在 `Office.onReady` 函数调用中，定位将单击处理程序分配到 `create-content-control` 按钮的行，并在该行后添加以下代码。

    ```js
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
    ```

1. 将以下函数添加到文件结尾。

    ```js
    async function replaceContentInControl() {
        await Word.run(async (context) => {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. 在 `replaceContentInControl()` 函数中，将 `TODO1` 替换为以下代码。注意：

    - `ContentControlCollection.getByTag` 方法将返回指定标记的所有内容控件的 `ContentControlCollection`。 我们使用 `getFirst` 来获取对所需控件的引用。

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

1. 验证是否已保存了对项目所做的所有更改。

### <a name="test-the-add-in"></a>测试加载项

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-word-start-server.md)]

1. 如果加载项任务窗格已在 Word 中打开，请转到“开始”选项卡并选择功能区中的“显示任务窗格”按钮以打开它。

1. 在任务窗格中，选择“**插入段落**”按钮，以确保文档顶部有包含“Microsoft 365”的段落。

1. 在文档中，选择文本“Microsoft 365”，然后选择 **创建内容控件** 按钮。观察此短语是否包装在标签为“服务名称”的标记中。

1. 选择“重命名服务”按钮，并观察内容控件的文本是否变成“Fabrikam Online Productivity Suite”。

    ![显示选择加载项按钮“创建内容控制”和“重命名服务”的结果屏幕截图。](../images/word-tutorial-content-control-2.png)

## <a name="next-steps"></a>后续步骤

在本教程中，你已创建 Word 任务窗格加载项，用于在 Word 文档中插入和替换文本、图像和其他内容。 若要了解有关构建 Word 加载项的详细信息，请继续阅读以下文章。

> [!div class="nextstepaction"]
> [Word 加载项概述](../word/word-add-ins-programming-overview.md)

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
