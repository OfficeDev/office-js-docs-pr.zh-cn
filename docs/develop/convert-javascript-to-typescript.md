---
title: 在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript
description: 了解如何将 Visual Studio 中的 Office 外接程序项目转换为使用 TypeScript。
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 1dbb3503a521f1a7c3e71764a50f02708b667a11
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719040"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="2d446-103">在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="2d446-103">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="2d446-104">可以使用 Visual Studio 中的 Office 加载项模板，创建使用 JavaScript 的加载项，再将加载项项目转换为使用 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="2d446-104">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="2d446-105">本文介绍了 Excel 加载项的此转换过程。</span><span class="sxs-lookup"><span data-stu-id="2d446-105">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="2d446-106">可以按照相同的过程操作，在 Visual Studio 中将其他类型的 Office 外接程序项目从 JavaScript 转换为 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="2d446-106">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="2d446-107">若不想使用 Visual Studio 创建 Office 加载项 TypeScript 项目，请按照任何 [5 分钟快速入门](../index.md)的“Yeoman 生成器”部分中的说明操作，并在[适用于 Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)显示提示时选择 `TypeScript`。</span><span class="sxs-lookup"><span data-stu-id="2d446-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](../index.md) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2d446-108">先决条件</span><span class="sxs-lookup"><span data-stu-id="2d446-108">Prerequisites</span></span>

- <span data-ttu-id="2d446-109">安装了 **Office/SharePoint 开发**工作负载的 [Visual Studio 2019](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="2d446-109">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="2d446-110">如果之前已安装 Visual Studio 2019，请[使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="2d446-110">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="2d446-111">如果尚未安装此工作负载，请使用 Visual Studio 安装程序进行[安装](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads)。</span><span class="sxs-lookup"><span data-stu-id="2d446-111">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span></span>

- <span data-ttu-id="2d446-112">TypeScript SDK 版本 2.3 或更高版本（适用于 Visual Studio 2019）</span><span class="sxs-lookup"><span data-stu-id="2d446-112">TypeScript SDK version 2.3 or later (for Visual Studio 2019)</span></span>

    > [!TIP]
    > <span data-ttu-id="2d446-113">在 [Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)中，选择“单个组件”\*\*\*\* 选项卡，然后向下滚动到“SDK、库和框架”\*\*\*\* 部分。</span><span class="sxs-lookup"><span data-stu-id="2d446-113">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="2d446-114">在该部分中，确保至少选择一个“TypeScript SDK”\*\*\*\* 组件（版本 2.3 或更高版本）。</span><span class="sxs-lookup"><span data-stu-id="2d446-114">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="2d446-115">如果一个“TypeScript SDK”\*\*\*\* 组件都没有选择，则选择最新可用版本的 SDK，然后选择“修改”\*\*\*\* 按钮以[安装该单个组件](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components)。</span><span class="sxs-lookup"><span data-stu-id="2d446-115">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span></span> 

- <span data-ttu-id="2d446-116">Excel 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="2d446-116">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="2d446-117">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="2d446-117">Create the add-in project</span></span>

1. <span data-ttu-id="2d446-118">在 Visual Studio 中，选择“**新建项目**”。</span><span class="sxs-lookup"><span data-stu-id="2d446-118">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="2d446-119">使用搜索框，输入“**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="2d446-119">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="2d446-120">选择“**Excel Web 加载项**”，然后选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="2d446-120">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="2d446-121">对项目命名，然后选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="2d446-121">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="2d446-122">在“创建 Office 加载项”\*\*\*\* 对话框窗口中，选择“将新功能添加到 Excel”\*\*\*\*，再选择“完成”\*\*\*\* 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="2d446-122">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="2d446-p105">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="2d446-p105">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="2d446-125">将加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="2d446-125">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="2d446-126">查找 **Home.js** 文件，并将其重命名为 **Home.ts**。</span><span class="sxs-lookup"><span data-stu-id="2d446-126">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="2d446-127">找到 **./Functions/FunctionFile.js** 文件，再将其重命名为 **FunctionFile.ts**。</span><span class="sxs-lookup"><span data-stu-id="2d446-127">Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.</span></span>

3. <span data-ttu-id="2d446-128">找到 **./Scripts/MessageBanner.js** 文件，再将其重命名为 **MessageBanner.ts**。</span><span class="sxs-lookup"><span data-stu-id="2d446-128">Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.</span></span>

4. <span data-ttu-id="2d446-129">从“**工具**”选项卡中，选择“**NuGet 程序包管理器**”，然后选择“**管理解决方案的 NuGet 程序包...**”。</span><span class="sxs-lookup"><span data-stu-id="2d446-129">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

5. <span data-ttu-id="2d446-130">在选中“**浏览**”选项卡的情况下，在搜索框中输入 **office-js.TypeScript.DefinitelyTyped**。</span><span class="sxs-lookup"><span data-stu-id="2d446-130">With the **Browse** tab selected, enter **office-js.TypeScript.DefinitelyTyped** into the search box.</span></span> <span data-ttu-id="2d446-131">安装或更新此程序包（如果已安装）。</span><span class="sxs-lookup"><span data-stu-id="2d446-131">Install or update this package if it is already installed.</span></span> <span data-ttu-id="2d446-132">这将把 Office.js 库的 TypeScript 类型定义添加到项目中。</span><span class="sxs-lookup"><span data-stu-id="2d446-132">This will add the TypeScript type definitions for the Office.js library to your project.</span></span>

6. <span data-ttu-id="2d446-133">在同一搜索框中，输入 **jquery.TypeScript.DefinitelyTyped**。</span><span class="sxs-lookup"><span data-stu-id="2d446-133">In the same search box, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="2d446-134">安装或更新此程序包（如果已安装）。</span><span class="sxs-lookup"><span data-stu-id="2d446-134">Install or update this package if it is already installed.</span></span> <span data-ttu-id="2d446-135">这将把 jQuery TypeScript 定义添加到项目中。</span><span class="sxs-lookup"><span data-stu-id="2d446-135">This will add the jQuery TypeScript definitions into your project.</span></span> <span data-ttu-id="2d446-136">jQuery 和 node.js 的程序包现在将显示在 Visual Studio 生成的称为 **packages.config** 的新文件中。</span><span class="sxs-lookup"><span data-stu-id="2d446-136">The packages for both jQuery and Office.js will now appear in a new file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="2d446-p108">在 TypeScript 项目中，可以混合使用 TypeScript 和 JavaScript 文件，项目都可以进行编译。这是因为 TypeScript 是键入的 JavaScript 超集，可以编译 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="2d446-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

7. <span data-ttu-id="2d446-139">在 **Home.ts** 中，找到行 `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` 并将其替换为以下内容：</span><span class="sxs-lookup"><span data-stu-id="2d446-139">In **Home.ts**, find the line `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` and replace it with the following:</span></span>

    ```TypeScript
    if(!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
    ```

    > [!NOTE]
    > <span data-ttu-id="2d446-140">目前，要使项目在转换为 TypeScript 后成功编译，你必须数值的形式指定要求集数量，如之前的代码片段中所示。</span><span class="sxs-lookup"><span data-stu-id="2d446-140">Currently, for the project to compile successfully after it's converted to TypeScript, you must specify the requirement set number as a numeric value as shown in the previous code snippet.</span></span> <span data-ttu-id="2d446-141">遗憾的是，这意味着你将无法使用 `isSetSupported` 来测试要求集 `1.10` 支持，因为数值 `1.10` 在运行时评估为 `1.1`。</span><span class="sxs-lookup"><span data-stu-id="2d446-141">Unfortunately this means you'll be unable to use `isSetSupported` to test for requirement set `1.10` support, as the numeric value `1.10` evaluates to `1.1` at runtime.</span></span> 
    > 
    > <span data-ttu-id="2d446-142">造成此问题的原因是 **office-js.TypeScript.DefinitelyTyped** NuGet 包当前已过时，这意味着你的项目无权访问 Office.js 的最新 TypeScript 定义。</span><span class="sxs-lookup"><span data-stu-id="2d446-142">This problem is due to the **office-js.TypeScript.DefinitelyTyped** NuGet package currently being outdated, which means your project doesn't have access to the latest TypeScript definitions for Office.js.</span></span> <span data-ttu-id="2d446-143">此问题正在解决中；问题解决后，本文将更新。</span><span class="sxs-lookup"><span data-stu-id="2d446-143">This issue is being addressed and this article will be updated when the issue is resolved.</span></span>

8. <span data-ttu-id="2d446-144">在 **Home.ts**中，找到 `Office.initialize = function (reason) {` 行并在其后面紧接着添加一行以填充全局 `window.Promise`，如下所示：</span><span class="sxs-lookup"><span data-stu-id="2d446-144">In **Home.ts**, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

9. <span data-ttu-id="2d446-145">在 **Home.ts** 中，找到 `displaySelectedCells` 函数，将整个函数替换为以下代码，然后保存文件：</span><span class="sxs-lookup"><span data-stu-id="2d446-145">In **Home.ts**, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

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

10. <span data-ttu-id="2d446-146">在 **./Scripts/MessageBanner.ts** 中，找到行 `_onResize(null);` 并将其替换为以下内容：</span><span class="sxs-lookup"><span data-stu-id="2d446-146">In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:</span></span>

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="2d446-147">运行转换后的外接程序项目</span><span class="sxs-lookup"><span data-stu-id="2d446-147">Run the converted add-in project</span></span>

1. <span data-ttu-id="2d446-p111">在 Visual Studio 中，按 **F5** 或选择“开始”\*\*\*\* 按钮以启动 Excel，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="2d446-p111">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="2d446-150">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="2d446-150">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="2d446-151">在工作表中，选择九个包含数字的单元格。</span><span class="sxs-lookup"><span data-stu-id="2d446-151">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="2d446-152">按任务窗格上的“突出显示”\*\*\*\* 按钮，以突出显示选定范围内所含数字最大的单元格。</span><span class="sxs-lookup"><span data-stu-id="2d446-152">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="2d446-153">Home.ts 代码文件</span><span class="sxs-lookup"><span data-stu-id="2d446-153">Home.ts code file</span></span>

<span data-ttu-id="2d446-p112">为方便参考，下面的代码片段展示了应用上述更改后的 **Home.ts** 文件内容。 此代码包括加载项运行至少所需的更改。</span><span class="sxs-lookup"><span data-stu-id="2d446-p112">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="2d446-156">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2d446-156">See also</span></span>

- [<span data-ttu-id="2d446-157">StackOverflow 上有关承诺实现的讨论</span><span class="sxs-lookup"><span data-stu-id="2d446-157">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [<span data-ttu-id="2d446-158">GitHub 上的 Office 外接程序示例</span><span class="sxs-lookup"><span data-stu-id="2d446-158">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
