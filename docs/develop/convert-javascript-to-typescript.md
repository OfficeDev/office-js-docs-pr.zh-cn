---
title: 在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript
description: 了解如何在 Office 中转换加载项Visual Studio TypeScript。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2771dbde52689200bc861aa208b33280776798b0
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148879"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript

可以使用 Visual Studio 中的 Office 加载项模板，创建使用 JavaScript 的加载项，再将加载项项目转换为使用 TypeScript。 本文介绍了 Excel 加载项的此转换过程。 可以按照相同的过程操作，在 Visual Studio 中将其他类型的 Office 外接程序项目从 JavaScript 转换为 TypeScript。

> [!IMPORTANT]
> 本文介绍确保按F5 时代码将转换为 JavaScript 所需的最少步骤，JavaScript 随后会自动旁加载至Office。 但是，代码不是非常"TypeScripty"。 例如，变量是使用关键字声明的，而不是使用指定类型 `var` `let` 声明的。 若要充分利用 TypeScript 的强键入，请考虑对代码进行进一步更改。 

> [!NOTE]
> 若不想使用 Visual Studio 创建 Office 加载项 TypeScript 项目，请按照任何 [5 分钟快速入门](../index.yml)的“Yeoman 生成器”部分中的说明操作，并在[适用于 Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)显示提示时选择 `TypeScript`。

## <a name="prerequisites"></a>先决条件

- 安装了 **Office/SharePoint 开发** 工作负载的 [Visual Studio 2019](https://www.visualstudio.com/vs/)

    > [!TIP]
    > 如果之前已安装 Visual Studio 2019，请 [使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发** 工作负载。 如果尚未安装此工作负载，请使用 Visual Studio 安装程序进行[安装](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-workloads)。

- TypeScript SDK 版本 2.3 或更高版本（适用于 Visual Studio 2019）

    > [!TIP]
    > 在 [Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)中，选择“单个组件”选项卡，然后向下滚动到“SDK、库和框架”部分。 在该部分中，确保至少选择一个“TypeScript SDK”组件（版本 2.3 或更高版本）。 如果未选择任何 **TypeScript SDK** 组件，请选择 SDK 的最新可用版本，然后选择"修改"[以安装该单个组件](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components)。

- Excel 2016 或更高版本

## <a name="create-the-add-in-project"></a>创建加载项项目

1. 在 Visual Studio 中，选择“**新建项目**”。

2. 使用搜索框，输入“**加载项**”。 选择“**Excel Web 加载项**”，然后选择“**下一步**”。

3. 对项目命名，然后选择“**创建**”。

4. 在“创建 Office 加载项”对话框窗口中，选择“将新功能添加到 Excel”，再选择“完成”以创建项目。

5. 此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”中。**Home.html** 文件在 Visual Studio 中打开。

## <a name="convert-the-add-in-project-to-typescript"></a>将加载项项目转换为使用 TypeScript

1. 查找 **Home.js** 文件，并将其重命名为 **Home.ts**。

2. 找到 **./Functions/FunctionFile.js** 文件，再将其重命名为 **FunctionFile.ts**。

3. 找到 **./Scripts/MessageBanner.js** 文件，再将其重命名为 **MessageBanner.ts**。

4. 从“**工具**”选项卡中，选择“**NuGet 程序包管理器**”，然后选择“**管理解决方案的 NuGet 程序包...**”。

5. 选中" **浏览"** 选项卡后，输入 **jquery。TypeScript.DefinitelyTyped**。 安装此程序包，或更新（如果已安装）。 这将确保在项目中包括 jQuery TypeScript 定义。 jQuery 包显示在由 Visual Studio 生成的文件中 **，packages.config。**

    > [!NOTE]
    > 在 TypeScript 项目中，可以混合使用 TypeScript 和 JavaScript 文件，项目都可以进行编译。这是因为 TypeScript 是键入的 JavaScript 超集，可以编译 JavaScript。

6. 在 **Home.ts** 中，找到 行，并紧接着添加一行以 `Office.initialize = function (reason) {` 填充全局 `window.Promise` ，如下所示。

    ```TypeScript
    Office.initialize = function (reason) {
        // Add the following line.
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

7. 在 **Home.ts** 中，找到 函数，将整个 `displaySelectedCells` 函数替换为以下代码，然后保存文件。

    ```TypeScript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }
    ```

8. 在 **./Scripts/MessageBanner.ts** 中，找到行 `_onResize(null);` 并将其替换为以下内容：

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a>运行转换后的外接程序项目

1. 在 Visual Studio 中，按 **F5** 或选择“开始”按钮以启动 Excel，功能区中显示有“显示任务窗格”加载项按钮。加载项本地托管在 IIS 上。

2. 在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。

3. 在工作表中，选择九个包含数字的单元格。

4. 按任务窗格上的“突出显示”按钮，以突出显示选定范围内所含数字最大的单元格。

## <a name="homets-code-file"></a>Home.ts 代码文件

为方便参考，下面的代码片段展示了应用上述更改后的 **Home.ts** 文件内容。 此代码包括加载项运行至少所需的更改。

```typescript
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If you're using Excel 2013, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");

            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(highlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function highlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
```

## <a name="see-also"></a>另请参阅

- [StackOverflow 上有关承诺实现的讨论](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [GitHub 上的 Office 外接程序示例](https://github.com/officedev)