---
title: 在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript
description: 了解如何将 Office 加载项项目转换为Visual Studio TypeScript。
ms.date: 09/01/2020
localization_priority: Normal
ms.openlocfilehash: 2134727a6065a1236dca313721d7721657e9a677
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839962"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="6ba82-103">在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="6ba82-103">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="6ba82-104">可以使用 Visual Studio 中的 Office 加载项模板，创建使用 JavaScript 的加载项，再将加载项项目转换为使用 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="6ba82-104">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="6ba82-105">本文介绍了 Excel 加载项的此转换过程。</span><span class="sxs-lookup"><span data-stu-id="6ba82-105">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="6ba82-106">可以按照相同的过程操作，在 Visual Studio 中将其他类型的 Office 外接程序项目从 JavaScript 转换为 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="6ba82-106">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6ba82-107">本文介绍了确保按F5 时代码将转换为 JavaScript（然后自动旁加载到 Office）所需的最少步骤。</span><span class="sxs-lookup"><span data-stu-id="6ba82-107">This article describes the *minimal* steps necessary to ensure that, when you press F5, the code will be transpiled to JavaScript which is then sideloaded automatically into Office.</span></span> <span data-ttu-id="6ba82-108">但是，代码不是"TypeScripty"。</span><span class="sxs-lookup"><span data-stu-id="6ba82-108">However, the code is not very "TypeScripty".</span></span> <span data-ttu-id="6ba82-109">例如，变量是使用关键字而不是关键字声明 `var` 的， `let` 并且它们不是使用指定类型声明的。</span><span class="sxs-lookup"><span data-stu-id="6ba82-109">For example, variables are declared with the `var` keyword instead of `let` and they are not declared with a specified type.</span></span> <span data-ttu-id="6ba82-110">若要充分利用 TypeScript 的强键入，请考虑进一步更改代码。</span><span class="sxs-lookup"><span data-stu-id="6ba82-110">To take full advantage of the strong typing of TypeScript, consider making further changes to the code.</span></span> 

> [!NOTE]
> <span data-ttu-id="6ba82-111">若不想使用 Visual Studio 创建 Office 加载项 TypeScript 项目，请按照任何 [5 分钟快速入门](../index.yml)的“Yeoman 生成器”部分中的说明操作，并在[适用于 Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)显示提示时选择 `TypeScript`。</span><span class="sxs-lookup"><span data-stu-id="6ba82-111">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6ba82-112">先决条件</span><span class="sxs-lookup"><span data-stu-id="6ba82-112">Prerequisites</span></span>

- <span data-ttu-id="6ba82-113">安装了 **Office/SharePoint 开发** 工作负载的 [Visual Studio 2019](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="6ba82-113">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="6ba82-114">如果之前已安装 Visual Studio 2019，请 [使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发** 工作负载。</span><span class="sxs-lookup"><span data-stu-id="6ba82-114">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="6ba82-115">如果尚未安装此工作负载，请使用 Visual Studio 安装程序进行[安装](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-workloads)。</span><span class="sxs-lookup"><span data-stu-id="6ba82-115">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-workloads).</span></span>

- <span data-ttu-id="6ba82-116">TypeScript SDK 版本 2.3 或更高版本（适用于 Visual Studio 2019）</span><span class="sxs-lookup"><span data-stu-id="6ba82-116">TypeScript SDK version 2.3 or later (for Visual Studio 2019)</span></span>

    > [!TIP]
    > <span data-ttu-id="6ba82-117">在 [Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)中，选择“单个组件”选项卡，然后向下滚动到“SDK、库和框架”部分。</span><span class="sxs-lookup"><span data-stu-id="6ba82-117">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="6ba82-118">在该部分中，确保至少选择一个“TypeScript SDK”组件（版本 2.3 或更高版本）。</span><span class="sxs-lookup"><span data-stu-id="6ba82-118">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="6ba82-119">如果一个“TypeScript SDK”组件都没有选择，则选择最新可用版本的 SDK，然后选择“修改”按钮以[安装该单个组件](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components)。</span><span class="sxs-lookup"><span data-stu-id="6ba82-119">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components).</span></span> 

- <span data-ttu-id="6ba82-120">Excel 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="6ba82-120">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="6ba82-121">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="6ba82-121">Create the add-in project</span></span>

1. <span data-ttu-id="6ba82-122">在 Visual Studio 中，选择“**新建项目**”。</span><span class="sxs-lookup"><span data-stu-id="6ba82-122">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="6ba82-123">使用搜索框，输入“**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="6ba82-123">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="6ba82-124">选择“**Excel Web 加载项**”，然后选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="6ba82-124">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="6ba82-125">对项目命名，然后选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="6ba82-125">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="6ba82-126">在“创建 Office 加载项”对话框窗口中，选择“将新功能添加到 Excel”，再选择“完成”以创建项目。</span><span class="sxs-lookup"><span data-stu-id="6ba82-126">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="6ba82-p106">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="6ba82-p106">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="6ba82-129">将加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="6ba82-129">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="6ba82-130">查找 **Home.js** 文件，并将其重命名为 **Home.ts**。</span><span class="sxs-lookup"><span data-stu-id="6ba82-130">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="6ba82-131">找到 **./Functions/FunctionFile.js** 文件，再将其重命名为 **FunctionFile.ts**。</span><span class="sxs-lookup"><span data-stu-id="6ba82-131">Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.</span></span>

3. <span data-ttu-id="6ba82-132">找到 **./Scripts/MessageBanner.js** 文件，再将其重命名为 **MessageBanner.ts**。</span><span class="sxs-lookup"><span data-stu-id="6ba82-132">Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.</span></span>

4. <span data-ttu-id="6ba82-133">从“**工具**”选项卡中，选择“**NuGet 程序包管理器**”，然后选择“**管理解决方案的 NuGet 程序包...**”。</span><span class="sxs-lookup"><span data-stu-id="6ba82-133">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

5. <span data-ttu-id="6ba82-134">选中 **"浏览"** 选项卡后，输入 **jquery。TypeScript.DefinitelyTyped**。</span><span class="sxs-lookup"><span data-stu-id="6ba82-134">With the **Browse** tab selected, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="6ba82-135">安装此程序包，或更新程序包（如果已安装）。</span><span class="sxs-lookup"><span data-stu-id="6ba82-135">Install this package, or update it if it's already installed.</span></span> <span data-ttu-id="6ba82-136">这将确保 jQuery TypeScript 定义包含在项目中。</span><span class="sxs-lookup"><span data-stu-id="6ba82-136">This will ensure the jQuery TypeScript definitions are included in your project.</span></span> <span data-ttu-id="6ba82-137">jQuery 的包显示在由 Visual Studio 生成的文件中 **，称为** packages.config。</span><span class="sxs-lookup"><span data-stu-id="6ba82-137">The packages for jQuery appear in a file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6ba82-p108">在 TypeScript 项目中，可以混合使用 TypeScript 和 JavaScript 文件，项目都可以进行编译。这是因为 TypeScript 是键入的 JavaScript 超集，可以编译 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="6ba82-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

6. <span data-ttu-id="6ba82-140">在 **Home.ts** 中，找到 `Office.initialize = function (reason) {` 行并在其后面紧接着添加一行以填充全局 `window.Promise`，如下所示：</span><span class="sxs-lookup"><span data-stu-id="6ba82-140">In **Home.ts**, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

7. <span data-ttu-id="6ba82-141">在 **Home.ts** 中，找到 `displaySelectedCells` 函数，将整个函数替换为以下代码，然后保存文件：</span><span class="sxs-lookup"><span data-stu-id="6ba82-141">In **Home.ts**, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

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

8. <span data-ttu-id="6ba82-142">在 **./Scripts/MessageBanner.ts** 中，找到行 `_onResize(null);` 并将其替换为以下内容：</span><span class="sxs-lookup"><span data-stu-id="6ba82-142">In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:</span></span>

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="6ba82-143">运行转换后的外接程序项目</span><span class="sxs-lookup"><span data-stu-id="6ba82-143">Run the converted add-in project</span></span>

1. <span data-ttu-id="6ba82-p109">在 Visual Studio 中，按 **F5** 或选择“开始”按钮以启动 Excel，功能区中显示有“显示任务窗格”加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="6ba82-p109">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="6ba82-146">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="6ba82-146">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="6ba82-147">在工作表中，选择九个包含数字的单元格。</span><span class="sxs-lookup"><span data-stu-id="6ba82-147">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="6ba82-148">按任务窗格上的“突出显示”按钮，以突出显示选定范围内所含数字最大的单元格。</span><span class="sxs-lookup"><span data-stu-id="6ba82-148">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="6ba82-149">Home.ts 代码文件</span><span class="sxs-lookup"><span data-stu-id="6ba82-149">Home.ts code file</span></span>

<span data-ttu-id="6ba82-p110">为方便参考，下面的代码片段展示了应用上述更改后的 **Home.ts** 文件内容。 此代码包括加载项运行至少所需的更改。</span><span class="sxs-lookup"><span data-stu-id="6ba82-p110">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="6ba82-152">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6ba82-152">See also</span></span>

- [<span data-ttu-id="6ba82-153">StackOverflow 上有关承诺实现的讨论</span><span class="sxs-lookup"><span data-stu-id="6ba82-153">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [<span data-ttu-id="6ba82-154">GitHub 上的 Office 外接程序示例</span><span class="sxs-lookup"><span data-stu-id="6ba82-154">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)