---
title: 在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 015fd9d7e9bf4412c09b76f0de5a97c9946e4d58
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016330"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript

可以使用 Visual Studio 中的 Office 加载项模板，创建使用 JavaScript 的加载项，再将加载项项目转换为使用 TypeScript。 使用 Visual Studio 创建加载项项目，无需从头开始创建 Office 加载项 TypeScript 项目。 

本文介绍了如何使用 Visual Studio 创建 Excel 加载项，再将加载项项目从使用 JavaScript 转换为使用 TypeScript。 可以按照相同的过程操作，在 Visual Studio 中将其他类型的 Office 加载项 JavaScript 项目转换为使用 TypeScript。

> [!NOTE]
> 若不想使用 Visual Studio 创建 Office 加载项 TypeScript 项目，请按照任何 [5 分钟快速入门](../index.yml) 的“任意编辑器”部分中的说明操作，并在 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office) 出现提示时选择 `TypeScript` 。

## <a name="prerequisites"></a>先决条件

- 安装了 **Office/SharePoint 开发**工作负载的 [Visual Studio 2017](https://www.visualstudio.com/vs/)

    > [!NOTE]
    > 如果之前已安装 Visual Studio 2017，请[使用 Visual Studio 安装程序](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发**工作负载。 

- TypeScript 2.3 for Visual Studio 2017

    > [!NOTE]
    > 虽然 TypeScript 应该会随 Visual Studio 2017 一起默认安装，但可以[使用 Visual Studio 安装程序](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)确认它是否已安装。 在 Visual Studio 安装程序中，选择“单个组件”**** 选项卡，再确认是否已在“SDK、库和框架”**** 下选中“TypeScript 2.3 SDK”****。

- Excel 2016 或更高版本

## <a name="create-the-add-in-project"></a>创建加载项项目

1. 打开 Visual Studio，在 Visual Studio 菜单栏中，依次选择“文件”**** > “新建”**** > “项目”****。

2. 在“Visual C#”**** 或“Visual Basic”**** 下的项目类型列表中，展开“Office/SharePoint”****，选择“加载项”****，再选择“Excel Web 加载项”**** 作为项目类型。 

3. 命名此项目，再选择“确定”****。

4. 在“创建 Office 加载项”**** 对话框窗口中，选择“将新功能添加到 Excel”****，再选择“完成”**** 以创建项目。

5. 此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”**** 中。**Home.html** 文件在 Visual Studio 中打开。

## <a name="convert-the-add-in-project-to-typescript"></a>将加载项项目转换为使用 TypeScript

1. 在“解决方案资源管理器”**** 中，将 **Home.js** 文件重命名为 **Home.ts**。

    > [!NOTE]
    > 在 TypeScript 项目中，可以混合使用 TypeScript 和 JavaScript 文件，项目都可以进行编译。这是因为 TypeScript 是键入的 JavaScript 超集，可以编译 JavaScript。 

2. 当出现提示时，选择“是”****，以确认要更改文件扩展名。

3. 在 Web 应用项目根目录中，新建 **Office.d.ts** 文件。

4. 在 Web 浏览器中，打开 [Office.js 的类型定义文件](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)。 将此文件的内容复制到剪贴板。

5. 在 Visual Studio 中，打开 **Office.d.ts** 文件，将剪贴板中的内容粘贴到此文件，并保存文件。

6. 在 Web 应用项目根目录中，新建 **jQuery.d.ts** 文件。

7. 在 Web 浏览器中，打开 [jQuery 的类型定义文件](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts)。 将此文件的内容复制到剪贴板。

8. 在 Visual Studio 中，打开 **jQuery.d.ts** 文件，将剪贴板中的内容粘贴到此文件，并保存文件。

9. 在 Visual Studio 中，转到 Web 应用项目根目录，新建 **tsconfig.json** 文件。

10. 打开 **tsconfig.json** 文件，将以下内容添加到此文件，并保存文件：

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. 打开 **Home.ts** 文件，并在文件顶部添加以下声明：

    ```javascript
    declare var fabric: any;
    ```

12. 在 **Home.ts** 文件中，将下面代码行中的 **'1.1'** 更改为 **1.1**（即删除引号），并保存文件：

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a>运行转换后的加载项项目

1. 在 Visual Studio 中，按 F5 或选择“开始”**** 按钮以启动 Excel，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。

2. 在 Excel 中，依次选择“开始”**** 选项卡和功能区中的“显示任务窗格”**** 按钮，打开加载项任务窗格。

3. 在工作表中，选择九个包含数字的单元格。

4. 按任务窗格上的“突出显示”**** 按钮，以突出显示选定范围内所含数字最大的单元格。

## <a name="homets-code-file"></a>Home.ts 代码文件

为方便参考，下面的代码片段展示了应用上述更改后的 **Home.ts** 文件内容。 此代码包括加载项运行至少所需的更改。

```javascript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016 or later, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
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
            $('#highlight-button').click(hightlightHighestValue);
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

    function hightlightHighestValue() {
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
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
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

* [StackOverflow 上有关承诺实现的讨论](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [GitHub 上的 Office 外接程序示例](https://github.com/officedev)
