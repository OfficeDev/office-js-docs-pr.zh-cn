---
title: '? Visual Studio ?? Office ?????????? TypeScript'
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 3d5b00d2a0b014dda350888030e5a50456292e1f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="c0bdb-102">? Visual Studio ?? Office ?????????? TypeScript</span><span class="sxs-lookup"><span data-stu-id="c0bdb-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="c0bdb-103">???? Visual Studio ?? Office ?????????? JavaScript ????????????????? TypeScript?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="c0bdb-104">?? Visual Studio ???????????????? Office ??? TypeScript ???</span><span class="sxs-lookup"><span data-stu-id="c0bdb-104">By using Visual Studio to create the add-in project, you avoid having to create your Office Add-in TypeScript project from scratch.</span></span> 

<span data-ttu-id="c0bdb-105">????????? Visual Studio ?? Excel ?????????????? JavaScript ????? TypeScript?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-105">This article shows you how to create an Excel add-in using Visual Studio and then convert the add-in project from JavaScript to TypeScript.</span></span> <span data-ttu-id="c0bdb-106">????????????? Visual Studio ??????? Office ??? JavaScript ??????? TypeScript?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-106">You can use the same process to convert other types of Office Add-in JavaScript projects to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="c0bdb-107">????? Visual Studio ?? Office ??? TypeScript ???????? [5 ??????](../index.yml)???????????????????[??? Office ???? Yeoman ???](https://github.com/OfficeDev/generator-office)??????? `TypeScript`?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quickstart](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c0bdb-108">????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-108">Prerequisites</span></span>

- <span data-ttu-id="c0bdb-109">??? **Office/SharePoint ??**????? [Visual Studio 2017](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="c0bdb-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="c0bdb-110">??????? Visual Studio 2017??[?? Visual Studio ????](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio)?????? **Office/SharePoint ??**?????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> 

- <span data-ttu-id="c0bdb-111">TypeScript 2.3 for Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="c0bdb-111">TypeScript 2.3 for Visual Studio 2017</span></span>

    > [!NOTE]
    > <span data-ttu-id="c0bdb-112">?? TypeScript ???? Visual Studio 2017 ??????????[?? Visual Studio ????](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio)?????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-112">TypeScript should be installed by default with Visual Studio 2017, but you can [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to confirm that it is installed.</span></span> <span data-ttu-id="c0bdb-113">? Visual Studio ??????????????****????????????SDK??????****????TypeScript 2.3 SDK?****?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-113">In the Visual Studio Installer, select the **Individual components** tab and then verify that **TypeScript 2.3 SDK** is selected under **SDKs, libraries, and frameworks**.</span></span>

- <span data-ttu-id="c0bdb-114">Excel 2016</span><span class="sxs-lookup"><span data-stu-id="c0bdb-114">Excel 2016</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="c0bdb-115">???????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-115">Create the add-in project</span></span>

1. <span data-ttu-id="c0bdb-116">?? Visual Studio?? Visual Studio ?????????????**** > ????**** > ????****?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-116">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="c0bdb-117">??Visual C#?****??Visual Basic?****?????????????Office/SharePoint?****????????****?????Excel Web ????****???????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-117">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="c0bdb-118">?????????????****?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-118">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="c0bdb-119">???? Office ????****????????????????? Excel?****????????****??????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-119">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="c0bdb-p104">???Visual Studio ????????????????????????????****??**Home.html** ??? Visual Studio ????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="c0bdb-122">??????????? TypeScript</span><span class="sxs-lookup"><span data-stu-id="c0bdb-122">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="c0bdb-123">????????????****??? **Home.js** ?????? **Home.ts**?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-123">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c0bdb-p105">? TypeScript ?????????? TypeScript ? JavaScript ????????????????? TypeScript ???? JavaScript ??????? JavaScript?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="c0bdb-126">????????????****?????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-126">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="c0bdb-127">? Web ??????????? **Office.d.ts** ???</span><span class="sxs-lookup"><span data-stu-id="c0bdb-127">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="c0bdb-128">? Web ??????? [Office.js ???????](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts)?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-128">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="c0bdb-129">??????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-129">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="c0bdb-130">? Visual Studio ???? **Office.d.ts** ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-130">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="c0bdb-131">? Web ??????????? **jQuery.d.ts** ???</span><span class="sxs-lookup"><span data-stu-id="c0bdb-131">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="c0bdb-132">? Web ??????? [jQuery ???????](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts)?</span><span class="sxs-lookup"><span data-stu-id="c0bdb-132">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span></span> <span data-ttu-id="c0bdb-133">??????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-133">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="c0bdb-134">? Visual Studio ???? **jQuery.d.ts** ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-134">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="c0bdb-135">? Visual Studio ???? Web ?????????? **tsconfig.json** ???</span><span class="sxs-lookup"><span data-stu-id="c0bdb-135">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="c0bdb-136">?? **tsconfig.json** ?????????????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-136">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="c0bdb-137">?? **Home.ts** ????????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-137">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```javascript
    declare var fabric: any;
    ```

12. <span data-ttu-id="c0bdb-138">? **Home.ts** ???????????? **'1.1'** ??? **1.1**??????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-138">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line, and save the file:</span></span>

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="c0bdb-139">???????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-139">Run the converted add-in project</span></span>

1. <span data-ttu-id="c0bdb-p108">? Visual Studio ??? F5 ???????****????? Excel????????????????****?????????????? IIS ??</span><span class="sxs-lookup"><span data-stu-id="c0bdb-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="c0bdb-142">? Excel ??????????****?????????????????****?????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-142">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="c0bdb-143">???????????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-143">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="c0bdb-144">?????????????****????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-144">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="c0bdb-145">Home.ts ????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-145">Home.ts code file</span></span>

<span data-ttu-id="c0bdb-146">???????????????????????? **Home.ts** ?????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-146">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="c0bdb-147">??????????????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-147">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c0bdb-148">????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-148">See also</span></span>

* [<span data-ttu-id="c0bdb-149">StackOverflow ??????????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-149">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="c0bdb-150">GitHub ?? Office ??????</span><span class="sxs-lookup"><span data-stu-id="c0bdb-150">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
