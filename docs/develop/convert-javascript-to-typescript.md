---
title: 在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript
description: ''
ms.date: 10/30/2018
ms.openlocfilehash: 9ea1cf421ce94d7756595950604ab3279e049c95
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724849"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="8d7f1-102">在 Visual Studio 中将 Office 加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="8d7f1-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="8d7f1-103">可以使用 Visual Studio 中的 Office 加载项模板，创建使用 JavaScript 的加载项，再将加载项项目转换为使用 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="8d7f1-104">本文介绍了 Excel 加载项的此转换过程。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-104">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="8d7f1-105">可以按照相同的过程操作，在 Visual Studio 中将其他类型的 Office 外接程序项目从 JavaScript 转换为 TypeScript。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-105">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="8d7f1-106">若不想使用 Visual Studio 创建 Office 外接程序 TypeScript 项目，请按照任何 [5 分钟快速入门](../index.yml)的“任意编辑器”部分中的说明操作，并在[适用于 Office 外接程序的 Yeoman 生成器](https://github.com/officedev/generator-office)显示提示时选择 `TypeScript`。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-106">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quick start](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="8d7f1-107">先决条件</span><span class="sxs-lookup"><span data-stu-id="8d7f1-107">Prerequisites</span></span>

- <span data-ttu-id="8d7f1-108">安装了 **Office/SharePoint 开发**工作负载的 [Visual Studio 2017](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="8d7f1-108">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="8d7f1-109">如果之前已安装 Visual Studio 2017，请[使用 Visual Studio 安装程序](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-109">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="8d7f1-110">如果尚未安装此工作负载，请使用 Visual Studio 安装程序进行[安装](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads)。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-110">If this workload is not yet installed, use the Visual Studio Installer to [install it](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).</span></span>

- <span data-ttu-id="8d7f1-111">TypeScript SDK 版本 2.3 或更高版本（适用于 Visual Studio 2017）</span><span class="sxs-lookup"><span data-stu-id="8d7f1-111">TypeScript SDK version 2.3 or later (for Visual Studio 2017)</span></span>

    > [!TIP]
    > <span data-ttu-id="8d7f1-112">在 [Visual Studio 安装程序](https://docs.microsoft.com/visualstudio/install/modify-visual-studio)中，选择“单个组件”\*\*\*\* 选项卡，然后向下滚动到“SDK、库和框架”\*\*\*\* 部分。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-112">In the [Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="8d7f1-113">在该部分中，确保至少选择一个“TypeScript SDK”\*\*\*\* 组件（版本 2.3 或更高版本）。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-113">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="8d7f1-114">如果一个“TypeScript SDK”\*\*\*\* 组件都没有选择，则选择最新可用版本的 SDK，然后选择“修改”\*\*\*\* 按钮以[安装该单个组件](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components)。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-114">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components).</span></span> 

- <span data-ttu-id="8d7f1-115">Excel 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="8d7f1-115">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="8d7f1-116">创建外接程序项目</span><span class="sxs-lookup"><span data-stu-id="8d7f1-116">Create the add-in project</span></span>

1. <span data-ttu-id="8d7f1-117">打开 Visual Studio，在 Visual Studio 菜单栏中，依次选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-117">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="8d7f1-118">在“Visual C#”\*\*\*\* 或“Visual Basic”\*\*\*\* 下的项目类型列表中，展开“Office/SharePoint”\*\*\*\*，选择“加载项”\*\*\*\*，再选择“Excel Web 加载项”\*\*\*\* 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-118">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="8d7f1-119">命名此项目，再选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-119">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="8d7f1-120">在“创建 Office 加载项”\*\*\*\* 对话框窗口中，选择“将新功能添加到 Excel”\*\*\*\*，再选择“完成”\*\*\*\* 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-120">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="8d7f1-p104">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="8d7f1-123">将加载项项目转换为使用 TypeScript</span><span class="sxs-lookup"><span data-stu-id="8d7f1-123">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="8d7f1-124">在“解决方案资源管理器”\*\*\*\* 中，将 **Home.js** 文件重命名为 **Home.ts**。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-124">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8d7f1-p105">在 TypeScript 项目中，可以混合使用 TypeScript 和 JavaScript 文件，项目都可以进行编译。这是因为 TypeScript 是键入的 JavaScript 超集，可以编译 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="8d7f1-127">当出现提示时，选择“是”\*\*\*\*，以确认要更改文件扩展名。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-127">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="8d7f1-128">在 Web 应用项目根目录中，新建 **Office.d.ts** 文件。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-128">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="8d7f1-129">在 Web 浏览器中，打开 [Office.js 的类型定义文件](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-129">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="8d7f1-130">将此文件的内容复制到剪贴板。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-130">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="8d7f1-131">在 Visual Studio 中，打开 **Office.d.ts** 文件，将剪贴板中的内容粘贴到此文件，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-131">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="8d7f1-132">在 Web 应用项目根目录中，新建 **jQuery.d.ts** 文件。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-132">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="8d7f1-133">在 Web 浏览器中，打开 [jQuery 的类型定义文件](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-133">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts).</span></span> <span data-ttu-id="8d7f1-134">将此文件的内容复制到剪贴板。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-134">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="8d7f1-135">在 Visual Studio 中，打开 **jQuery.d.ts** 文件，将剪贴板中的内容粘贴到此文件，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-135">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="8d7f1-136">在 Visual Studio 中，转到 Web 应用项目根目录，新建 **tsconfig.json** 文件。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-136">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="8d7f1-137">打开 **tsconfig.json** 文件，将以下内容添加到此文件，并保存文件：</span><span class="sxs-lookup"><span data-stu-id="8d7f1-137">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```json
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="8d7f1-138">打开“Home.ts”\*\*\*\* 文件，并在文件顶部添加以下声明：</span><span class="sxs-lookup"><span data-stu-id="8d7f1-138">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```typescript
    declare var fabric: any;
    ```

12. <span data-ttu-id="8d7f1-139">在“Home.ts”\*\*\*\* 文件中，将下面行中的“'1.1'”\*\*\*\* 更改为“1.1”\*\*\*\*（即删除引号）：</span><span class="sxs-lookup"><span data-stu-id="8d7f1-139">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line:</span></span>

    ```typescript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

13. <span data-ttu-id="8d7f1-140">在“Home.ts”\*\*\*\* 文件中，找到 `displaySelectedCells` 函数，将整个函数替换为以下代码，并保存该文件：</span><span class="sxs-lookup"><span data-stu-id="8d7f1-140">In the **Home.ts** file, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

    ```typescript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
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

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="8d7f1-141">运行转换后的外接程序项目</span><span class="sxs-lookup"><span data-stu-id="8d7f1-141">Run the converted add-in project</span></span>

1. <span data-ttu-id="8d7f1-p108">在 Visual Studio 中，按 **F5** 或选择“开始”\*\*\*\* 按钮以启动 Excel，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="8d7f1-144">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="8d7f1-145">在工作表中，选择九个包含数字的单元格。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-145">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="8d7f1-146">按任务窗格上的“突出显示”\*\*\*\* 按钮，以突出显示选定范围内所含数字最大的单元格。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-146">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="8d7f1-147">Home.ts 代码文件</span><span class="sxs-lookup"><span data-stu-id="8d7f1-147">Home.ts code file</span></span>

<span data-ttu-id="8d7f1-148">为方便参考，下面的代码片段展示了应用上述更改后的 **Home.ts** 文件内容。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-148">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="8d7f1-149">此代码包括加载项运行至少所需的更改。</span><span class="sxs-lookup"><span data-stu-id="8d7f1-149">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```typescript
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

## <a name="see-also"></a><span data-ttu-id="8d7f1-150">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8d7f1-150">See also</span></span>

* [<span data-ttu-id="8d7f1-151">StackOverflow 上有关承诺实现的讨论</span><span class="sxs-lookup"><span data-stu-id="8d7f1-151">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="8d7f1-152">GitHub 上的 Office 外接程序示例</span><span class="sxs-lookup"><span data-stu-id="8d7f1-152">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
