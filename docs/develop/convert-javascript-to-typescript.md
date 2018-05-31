---
title: 在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 3d5b00d2a0b014dda350888030e5a50456292e1f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437281"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="5c571-102">在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="5c571-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="5c571-103">可以使用 Visual Studio 中的 Office 加载项模板，创建使用 JavaScript 的加载项，再将加载项项目转换为使用 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="5c571-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="5c571-104">使用 Visual Studio 创建加载项项目，无需从头开始创建 Office 加载项 TypeScript 项目。</span><span class="sxs-lookup"><span data-stu-id="5c571-104">By using Visual Studio to create the add-in project, you avoid having to create your Office Add-in TypeScript project from scratch.</span></span> 

<span data-ttu-id="5c571-105">本文介绍了如何使用 Visual Studio 创建 Excel 加载项，再将加载项项目从使用 JavaScript 转换为使用 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="5c571-105">This article shows you how to create an Excel add-in using Visual Studio and then convert the add-in project from JavaScript to TypeScript.</span></span> <span data-ttu-id="5c571-106">可以按照相同的过程操作，在 Visual Studio 中将其他类型的 Office 加载项 JavaScript 项目转换为使用 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="5c571-106">You can use the same process to convert other types of Office Add-in JavaScript projects to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="5c571-107">若不想使用 Visual Studio 创建 Office 加载项 TypeScript 项目，请按照任何 [5 分钟快速入门](../index.yml)的“任意编辑器”部分中的说明操作，并在[适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)出现提示时选择 `TypeScript`。</span><span class="sxs-lookup"><span data-stu-id="5c571-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quickstart](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5c571-108">先决条件</span><span class="sxs-lookup"><span data-stu-id="5c571-108">Prerequisites</span></span>

- <span data-ttu-id="5c571-109">安装了 **Office/SharePoint 开发**工作负载的 [Visual Studio 2017](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="5c571-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="5c571-110">如果之前已安装 Visual Studio 2017，请[使用 Visual Studio 安装程序](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="5c571-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> 

- <span data-ttu-id="5c571-111">TypeScript 2.3 for Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="5c571-111">TypeScript 2.3 for Visual Studio 2017</span></span>

    > [!NOTE]
    > <span data-ttu-id="5c571-112">虽然 TypeScript 应该会随 Visual Studio 2017 一起默认安装，但可以[使用 Visual Studio 安装程序](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio)确认它是否已安装。</span><span class="sxs-lookup"><span data-stu-id="5c571-112">TypeScript should be installed by default with Visual Studio 2017, but you can [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to confirm that it is installed.</span></span> <span data-ttu-id="5c571-113">在 Visual Studio 安装程序中，选择“单个组件”**** 选项卡，再确认是否已在“SDK、库和框架”**** 下选中“TypeScript 2.3 SDK”****。</span><span class="sxs-lookup"><span data-stu-id="5c571-113">In the Visual Studio Installer, select the **Individual components** tab and then verify that **TypeScript 2.3 SDK** is selected under **SDKs, libraries, and frameworks**.</span></span>

- <span data-ttu-id="5c571-114">Excel 2016</span><span class="sxs-lookup"><span data-stu-id="5c571-114">Excel 2016</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="5c571-115">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="5c571-115">Create the add-in project</span></span>

1. <span data-ttu-id="5c571-116">打开 Visual Studio，在 Visual Studio 菜单栏中，依次选择“文件”**** > “新建”**** > “项目”****。</span><span class="sxs-lookup"><span data-stu-id="5c571-116">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="5c571-117">在“Visual C#”**** 或“Visual Basic”**** 下的项目类型列表中，展开“Office/SharePoint”****，选择“加载项”****，再选择“Excel Web 加载项”**** 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="5c571-117">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="5c571-118">命名此项目，再选择“确定”****。</span><span class="sxs-lookup"><span data-stu-id="5c571-118">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="5c571-119">在“创建 Office 加载项”**** 对话框窗口中，选择“将新功能添加到 Excel”****，再选择“完成”**** 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="5c571-119">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="5c571-p104">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”**** 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="5c571-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="5c571-122">将加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="5c571-122">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="5c571-123">在“解决方案资源管理器”**** 中，将 **Home.js** 文件重命名为 **Home.ts**。</span><span class="sxs-lookup"><span data-stu-id="5c571-123">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5c571-p105">在 TypeScript 项目中，可以混合使用 TypeScript 和 JavaScript 文件，项目都可以进行编译。这是因为 TypeScript 是键入的 JavaScript 超集，可以编译 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="5c571-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="5c571-126">当出现提示时，选择“是”****，以确认要更改文件扩展名。</span><span class="sxs-lookup"><span data-stu-id="5c571-126">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="5c571-127">在 Web 应用项目根目录中，新建 **Office.d.ts** 文件。</span><span class="sxs-lookup"><span data-stu-id="5c571-127">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="5c571-128">在 Web 浏览器中，打开 [Office.js 的类型定义文件](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="5c571-128">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="5c571-129">将此文件的内容复制到剪贴板。</span><span class="sxs-lookup"><span data-stu-id="5c571-129">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="5c571-130">在 Visual Studio 中，打开 **Office.d.ts** 文件，将剪贴板中的内容粘贴到此文件，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="5c571-130">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="5c571-131">在 Web 应用项目根目录中，新建 **jQuery.d.ts** 文件。</span><span class="sxs-lookup"><span data-stu-id="5c571-131">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="5c571-132">在 Web 浏览器中，打开 [jQuery 的类型定义文件](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="5c571-132">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span></span> <span data-ttu-id="5c571-133">将此文件的内容复制到剪贴板。</span><span class="sxs-lookup"><span data-stu-id="5c571-133">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="5c571-134">在 Visual Studio 中，打开 **jQuery.d.ts** 文件，将剪贴板中的内容粘贴到此文件，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="5c571-134">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="5c571-135">在 Visual Studio 中，转到 Web 应用项目根目录，新建 **tsconfig.json** 文件。</span><span class="sxs-lookup"><span data-stu-id="5c571-135">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="5c571-136">打开 **tsconfig.json** 文件，将以下内容添加到此文件，并保存文件：</span><span class="sxs-lookup"><span data-stu-id="5c571-136">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="5c571-137">打开 **Home.ts** 文件，并在文件顶部添加以下声明：</span><span class="sxs-lookup"><span data-stu-id="5c571-137">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```javascript
    declare var fabric: any;
    ```

12. <span data-ttu-id="5c571-138">在 **Home.ts** 文件中，将下面代码行中的 **'1.1'** 更改为 **1.1**（即删除引号），并保存文件：</span><span class="sxs-lookup"><span data-stu-id="5c571-138">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line, and save the file:</span></span>

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="5c571-139">运行转换后的加载项项目</span><span class="sxs-lookup"><span data-stu-id="5c571-139">Run the converted add-in project</span></span>

1. <span data-ttu-id="5c571-p108">在 Visual Studio 中，按 F5 或选择“开始”**** 按钮以启动 Excel，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="5c571-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="5c571-142">在 Excel 中，依次选择“开始”**** 选项卡和功能区中的“显示任务窗格”**** 按钮，打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="5c571-142">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="5c571-143">在工作表中，选择九个包含数字的单元格。</span><span class="sxs-lookup"><span data-stu-id="5c571-143">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="5c571-144">按任务窗格上的“突出显示”**** 按钮，以突出显示选定范围内所含数字最大的单元格。</span><span class="sxs-lookup"><span data-stu-id="5c571-144">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="5c571-145">Home.ts 代码文件</span><span class="sxs-lookup"><span data-stu-id="5c571-145">Home.ts code file</span></span>

<span data-ttu-id="5c571-146">为方便参考，下面的代码片段展示了应用上述更改后的 **Home.ts** 文件内容。</span><span class="sxs-lookup"><span data-stu-id="5c571-146">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="5c571-147">此代码包括加载项运行至少所需的更改。</span><span class="sxs-lookup"><span data-stu-id="5c571-147">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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
            
            // If not using Excel 2016, use fallback logic.
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

## <a name="see-also"></a><span data-ttu-id="5c571-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5c571-148">See also</span></span>

* [<span data-ttu-id="5c571-149">StackOverflow 上有关承诺实现的讨论</span><span class="sxs-lookup"><span data-stu-id="5c571-149">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="5c571-150">GitHub 上的 Office 外接程序示例</span><span class="sxs-lookup"><span data-stu-id="5c571-150">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
