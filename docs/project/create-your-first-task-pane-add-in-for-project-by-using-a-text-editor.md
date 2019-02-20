---
title: 使用文本编辑器为 Microsoft Project 创建首个任务窗格加载项
description: ''
ms.date: 12/17/2018
localization_priority: Normal
ms.openlocfilehash: fb218dff1bed6b7723597a628db6217a5f149771
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/19/2019
ms.locfileid: "29389478"
---
# <a name="create-your-first-task-pane-add-in-for-microsoft-project-by-using-a-text-editor"></a><span data-ttu-id="e5014-102">使用文本编辑器为 Microsoft Project 创建首个任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="e5014-102">Create your first task pane add-in for Microsoft Project by using a text editor</span></span>

<span data-ttu-id="e5014-103">你可以使用适用于 Office 加载项的 Yeoman 生成器为 Project Standard 2013、Project Professional 2013 或更高版本创建任务窗格加载项。本文介绍如何创建一个简单的加载项，该加载项使用指向文件共享上的 HTML 文件的 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="e5014-103">You can create a task pane add-in for Project Standard 2013, Project Professional 2013, or later verions using the Yeoman generator for Office Add-ins. This article describes how to create a simple add-in that uses an XML manifest that points to an HTML file on a file share.</span></span> <span data-ttu-id="e5014-104">Project OM Test 示例加载项用于测试一些 JavaScript 功能，这些功能为加载项使用对象模型。使用 Project 中的“信任中心”\*\*\*\* 注册包含清单文件的文件共享后，你可以从功能区上的“Project”\*\*\*\* 选项卡打开任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="e5014-104">The Project OM Test sample add-in tests some JavaScript functions that use the object model for add-ins. After you use the  **Trust Center** in Project to register the file share that contains the manifest file, you can open the task pane add-in from the **Project** tab on the ribbon.</span></span> <span data-ttu-id="e5014-105">（本文中的示例代码基于 Microsoft Corporation 的 Arvind lyer 所做的测试应用程序。）</span><span class="sxs-lookup"><span data-stu-id="e5014-105">(The sample code in this article is based on a test application by Arvind Iyer, Microsoft Corporation.)</span></span>

<span data-ttu-id="e5014-106">Project 与其他 Microsoft Office 客户端使用相同的加载项清单架构以及许多相同的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="e5014-106">Project uses the same add-in manifest schema that other Microsoft Office clients use, and much of the same JavaScript API.</span></span> <span data-ttu-id="e5014-107">Project 2013 SDK 下载的 `Samples\Apps` 子目录中提供了本文所述的加载项的完整代码。</span><span class="sxs-lookup"><span data-stu-id="e5014-107">The complete code for the add-in that is described in this article is available in the  `Samples\Apps` subdirectory of the Project 2013 SDK download.</span></span>

<span data-ttu-id="e5014-108">“Project OM 测试”示例加载项可以获取任务的 GUID，以及应用和有效项目的属性。</span><span class="sxs-lookup"><span data-stu-id="e5014-108">The Project OM Test sample add-in can get the GUID of a task and properties of the application and the active project.</span></span> <span data-ttu-id="e5014-109">如果 Project Professional 2013 打开 SharePoint 库中的项目，加载项可以显示项目的 URL。</span><span class="sxs-lookup"><span data-stu-id="e5014-109">If Project Professional 2013 opens a project that is in a SharePoint library, the add-in can show the URL of the project.</span></span> 

<span data-ttu-id="e5014-p104">[Project 2013 SDK 下载内容](https://www.microsoft.com/download/details.aspx?id=30435%20)包含完整源代码。提取和安装 Project2013SDK.msi 文件中的 SDK 和示例时，请在 `\Samples\Apps\Copy_to_AppManifests_FileShare` 子目录中查找清单文件，并在 `\Samples\Apps\Copy_to_AppSource_FileShare` 子目录中查找源代码。</span><span class="sxs-lookup"><span data-stu-id="e5014-p104">The [Project 2013 SDK download](https://www.microsoft.com/download/details.aspx?id=30435%20) includes the complete source code. When you extract and install the SDK and samples that are in the Project2013SDK.msi file, see the `\Samples\Apps\Copy_to_AppManifests_FileShare` subdirectory for the manifest file and the `\Samples\Apps\Copy_to_AppSource_FileShare` subdirectory for the source code.</span></span> 

<span data-ttu-id="e5014-112">JSOMCall.html 示例使用 office.js 文件和 project-15.js 文件中包含的 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="e5014-112">The JSOMCall.html sample uses JavaScript functions in the office.js file and project-15.js file, which are included.</span></span> <span data-ttu-id="e5014-113">可以使用相应的调试文件（office.debug.js 和 project-15.debug.js）检查这些函数。</span><span class="sxs-lookup"><span data-stu-id="e5014-113">You can use the corresponding debug files (office.debug.js and project-15.debug.js) to examine the functions.</span></span>

<span data-ttu-id="e5014-114">若要了解如何在 Office 加载项中使用 JavaScript，请参阅[了解适用于 Office 的 JavaScript API](../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="e5014-114">For an introduction to using JavaScript in Office Add-ins, see [Understanding the JavaScript API for Office](../develop/understanding-the-javascript-api-for-office.md).</span></span>

## <a name="procedure-1-to-create-the-add-in-manifest-file"></a><span data-ttu-id="e5014-p106">过程 1：创建加载项清单文件</span><span class="sxs-lookup"><span data-stu-id="e5014-p106">Procedure 1. To create the add-in manifest file</span></span>

<span data-ttu-id="e5014-p107">在本地目录中创建 XML 文件。此 XML 文件包括 [Office 加载项 XML 清单](../develop/add-in-manifests.md)中描述的 **OfficeApp** 元素和子元素。例如，创建包含以下 XML（更改 **Id** 元素的 GUID 值）的 JSOM_SimpleOMCalls.xml 文件。</span><span class="sxs-lookup"><span data-stu-id="e5014-p107">Create an XML file in a local directory. The XML file includes the **OfficeApp** element and child elements, which are described in the [Office Add-ins XML manifest](../develop/add-in-manifests.md). For example, create a file named JSOM_SimpleOMCalls.xml that contains the following XML (change the GUID value of the **Id** element).</span></span>
    
```XML
<?xml version="1.0" encoding="utf-8"?>
   <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
              xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
              xsi:type="TaskPaneApp">
     <Id>93A26520-9414-492F-994B-4983A1C7A607</Id>
     <Version>15.0</Version>
     <ProviderName>Microsoft</ProviderName>
     <DefaultLocale>en-us</DefaultLocale>
     <DisplayName DefaultValue="Project OM Test">
       <Override Locale="fr-fr" Value="Le Project OM Test"/>
     </DisplayName>
     <Description DefaultValue="Test the task pane add-in object model for Project - English (US)">
       <Override Locale="fr-fr" Value="Test the task pane add-in object model for Project - French (France)"/>
     </Description>
     <Hosts>
       <Host Name="Project"/>
       <Host Name="Workbook"/>
       <Host Name="Document"/>
     </Hosts>
    <DefaultSettings>
       <SourceLocation DefaultValue="\\ServerName\AppSource\JSOMCall.html">
         <Override Locale="fr-fr" Value="\\ServerName\AppSource\JSOMCall.html"/>
       </SourceLocation>
     </DefaultSettings>
     <Permissions>ReadWriteDocument</Permissions>
     <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
       <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
     </IconUrl>
     <AllowSnapshot>true</AllowSnapshot>
   </OfficeApp>
```

<span data-ttu-id="e5014-p108">对于 Project，**OfficeApp** 元素必须包括 `xsi:type="TaskPaneApp"` 属性值。**Id** 元素是 GUID。**SourceLocation** 值必须是加载项 HTML 源文件或任务窗格中运行的 Web 应用的文件共享路径或 SharePoint URL。有关清单文件中其他元素的解释，请参阅 [Project 任务窗格加载项](../project/project-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="e5014-p108">For Project, the **OfficeApp** element must include the `xsi:type="TaskPaneApp"` attribute value. The **Id** element is a GUID. The **SourceLocation** value must be a file share path or a SharePoint URL for the add-in HTML source file or the web application that runs in the task pane. For an explanation of the other elements in manifest file, see [Task pane add-ins for Project](../project/project-add-ins.md).</span></span>
    
<span data-ttu-id="e5014-p109">过程 2 演示如何创建 JSOM_SimpleOMCalls.XML 清单为 Project 测试加载项指定的 HTML 文件。HTML 文件中指定的按钮调用相关 JavaScript 函数。可以在 HTML 文件内添加 JavaScript 函数，或将它们放在一个单独的 .js 文件中。</span><span class="sxs-lookup"><span data-stu-id="e5014-p109">Procedure 2 shows how to create the HTML file that the JSOM_SimpleOMCalls.xml manifest specifies for the Project test add-in. Buttons that are specified in the HTML file call related JavaScript functions. You can add the JavaScript functions within the HTML file, or put them in a separate .js file.</span></span>

## <a name="procedure-2-to-create-the-source-files-for-the-project-om-test-add-in"></a><span data-ttu-id="e5014-p110">过程 2：创建“Project OM 测试”加载项的源文件</span><span class="sxs-lookup"><span data-stu-id="e5014-p110">Procedure 2. To create the source files for the Project OM Test add-in</span></span>

1. <span data-ttu-id="e5014-129">创建 HTML 文件，采用 JSOM_SimpleOMCalls.xml 清单中 **SourceLocation** 元素指定的名称。</span><span class="sxs-lookup"><span data-stu-id="e5014-129">Create an HTML file with a name that is specified by the **SourceLocation** element in the JSOM_SimpleOMCalls.xml manifest.</span></span> 

   <span data-ttu-id="e5014-130">例如，在 `C:\Project\AppSource` 目录中创建 theJSOMCall.html 文件。</span><span class="sxs-lookup"><span data-stu-id="e5014-130">For example, create theJSOMCall.html file in the `C:\Project\AppSource` directory.</span></span> <span data-ttu-id="e5014-131">可以使用简单的文本编辑器创建源文件，但是使用诸如 Visual Studio Code 的工具可以使操作更为简单，这适用于特定文档类型（如 HTML 和 JavaScript），并具有其他编辑辅助功能。</span><span class="sxs-lookup"><span data-stu-id="e5014-131">Although you can use a simple text editor to create the source files, it is easier to use a tool such as Visual Studio code, which works with specific document types (such as HTML and JavaScript) and has other editing aids.</span></span> <span data-ttu-id="e5014-132">如果还未执行 [Project 任务窗格加载项](../project/project-add-ins.md)所述的必应搜索示例，过程 3 将演示如何创建清单指定的 `\\ServerName\AppSource` 文件共享。</span><span class="sxs-lookup"><span data-stu-id="e5014-132">If you have not already done the Bing Search example that is described in [Task pane add-ins for Project](../project/project-add-ins.md), Procedure 3 shows how to create the `\\ServerName\AppSource` file share that the manifest specifies.</span></span>
    
   <span data-ttu-id="e5014-133">JSOMCall.html 文件在 Microsoft Office 2013 应用程序中使用通用 MicrosoftAjax.js 文件实现 AJAX 功能并使用 Office.js 文件实现外接程序功能。</span><span class="sxs-lookup"><span data-stu-id="e5014-133">The JSOMCall.html file uses the common MicrosoftAjax.js file for AJAX functionality and the Office.js file for the add-in functionality in Microsoft Office 2013 applications.</span></span>

    ```HTML
    <!DOCTYPE html>
    <html>
        <head>
            <title>Project OM Sample Code</title>
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <script type="text/javascript" src="MicrosoftAjax.js"></script>

            <!-- Use the CDN reference to office.js when deploying your add-in. -->
            <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script> -->
            <script type="text/javascript" src="Office.js"></script>
            <script type="text/javascript" src="JSOM_Sample.js"></script>
        </head>
        <body>
            <div id="Common_JSOM_API">
                OBJECT MODEL TESTS
            </div>

            <textarea id="text" rows="6" cols="25">This is the text result.</textarea>
        </body>
    </html>
    ```

   <span data-ttu-id="e5014-134">**textarea** 元素指定显示 JavaScript 函数结果的文本框。</span><span class="sxs-lookup"><span data-stu-id="e5014-134">The **textarea** element specifies a text box that shows results of the JavaScript functions.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="e5014-135">为了让“Project OM 测试”示例能够正常运行，请将 Project 2013 SDK 下载内容中的下列文件复制到 JSOMCall.html 文件所在的相同目录：Office.js、Project-15.js 和 MicrosoftAjax.js。</span><span class="sxs-lookup"><span data-stu-id="e5014-135">For the Project OM Test sample to work, copy the following files from the Project 2013 SDK download to the same directory as the JSOMCall.html file: Office.js, Project-15.js, and MicrosoftAjax.js.</span></span>

   <span data-ttu-id="e5014-p112">第 2 步为“Project OM 测试”示例加载项使用的特定函数添加 JSOM_Sample.js 文件。在后续步骤中，将为调用 JavaScript 函数的按钮添加其他 HTML 元素。</span><span class="sxs-lookup"><span data-stu-id="e5014-p112">Step 2 adds the JSOM_Sample.js file for specific functions that the Project OM Test sample add-in uses. In later steps, you will add other HTML elements for buttons that call JavaScript functions.</span></span>
    
2. <span data-ttu-id="e5014-138">在 JSOMCall.html 文件所在的相同目录中，创建 JavaScript 文件 JSOM_Sample.js。</span><span class="sxs-lookup"><span data-stu-id="e5014-138">Create a JavaScript file named JSOM_Sample.js in the same directory as the JSOMCall.html file.</span></span> 

   <span data-ttu-id="e5014-p113">下面的代码使用 Office.js 文件中的函数，获取应用上下文和文档信息。**text** 对象是 HTML 文件中 \*\* textarea\*\* 控件的 ID。</span><span class="sxs-lookup"><span data-stu-id="e5014-p113">The following code gets the application context and document information by using functions in the Office.js file. The **text** object is the ID of the **textarea** control in the HTML file.</span></span>
    
   <span data-ttu-id="e5014-p114">**\_projDoc** 变量是使用 **ProjectDocument** 对象进行初始化。代码包含一些简单的错误处理函数，以及获取应用上下文和项目文档上下文属性的 **getContextValues** 函数。若要详细了解 Project 的 JavaScript 对象模型，请参阅[适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="e5014-p114">The **\_projDoc** variable is initialized with a **ProjectDocument** object. The code includes some simple error handling functions, and the **getContextValues** function that gets application context and project document context properties. For more information about the JavaScript object model for Project, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office).</span></span>

    ```javascript
    /*
    * JavaScript functions for the Project OM Test example app
    * in the Project 2013 SDK.
    */

    var _projDoc;
    var _app;
    var taskGuid;
    var resourceGuid;

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            _projDoc = Office.context.document;
            _app = Office.context;
        });
    }

    function logError(errorText) {
        text.value = "Error in " + errorText;
    }

    function logEventError(erroneousEvent) {
        logError("event " + erroneousEvent);
    }

    function logMethodError(methodName, errorName, errorMessage) {
        logError(methodName + " method.\nError name: " + errorName + "\nMessage: " + errorMessage);
    }

    // . . . Add other JavaScript functions here.

    function getContextValues() {
        getDocumentUrl();
        getDocumentMode();
        getApplicationContentLanguage();
        getApplicationDisplayLanguage();
    }

    function getDocumentUrl() {
        text.value ="Document URL:\n" +_projDoc.url;
    }

    function getDocumentMode() {
        var docMode = _projDoc.mode;
        text.value = text.value + "\n\nDocument mode: " + docMode;
    }

    function getApplicationContentLanguage() {
        text.value = text.value + "\nApp language: " + _app.contentLanguage;
    }

    function getApplicationDisplayLanguage() {
        text.value = text.value + "\nDisplay language: " + _app.displayLanguage;
    }
    ```

   <span data-ttu-id="e5014-p115">有关 Office.debug.js 文件中函数的信息，请参见 [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)。例如，**getDocumentUrl** 函数获取打开的项目的 URL 或文件路径。</span><span class="sxs-lookup"><span data-stu-id="e5014-p115">For information about the functions in the Office.debug.js file, see [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office). For example, the **getDocumentUrl** function gets the URL or file path of the open project.</span></span>
    
3. <span data-ttu-id="e5014-146">添加调用 Office.js 和 Project-15.js 中异步函数的 JavaScript 函数，以获取选定数据：</span><span class="sxs-lookup"><span data-stu-id="e5014-146">Add JavaScript functions that call asynchronous functions in Office.js and Project-15.js to get selected data:</span></span>
    
   - <span data-ttu-id="e5014-p116">例如，**getSelectedDataAsync** 是 Office.js 中的常规函数，用于获取选定数据的无格式文本。有关详细信息，请参阅 [AsyncResult 对象](https://docs.microsoft.com/javascript/api/office/office.asyncresult)。</span><span class="sxs-lookup"><span data-stu-id="e5014-p116">For example, **getSelectedDataAsync** is a general function in Office.js that gets unformatted text for the selected data. For more information, see [AsyncResult object](https://docs.microsoft.com/javascript/api/office/office.asyncresult).</span></span>
    
   - <span data-ttu-id="e5014-p117">Project-15.js 中的 **getSelectedTaskAsync** 函数用于获取选定任务的 GUID。同样，**getSelectedResourceAsync** 函数用于获取选定资源的 GUID。如果在未选择任务或资源时调用这些函数，函数会显示未定义错误。</span><span class="sxs-lookup"><span data-stu-id="e5014-p117">The **getSelectedTaskAsync** function in Project-15.js gets the GUID of the selected task. Similarly, the **getSelectedResourceAsync** function gets the GUID of the selected resource. If you call those functions when a task or a resource is not selected, the functions show an undefined error.</span></span>
    
   - <span data-ttu-id="e5014-p118">**getTaskAsync** 函数用于获取任务名称和已分配资源的名称。如果任务位于同步的 SharePoint 任务列表中，**getTaskAsync** 可获取 SharePoint 列表中的任务 ID；否则，SharePoint 任务 ID 为 0。</span><span class="sxs-lookup"><span data-stu-id="e5014-p118">The **getTaskAsync** function gets the task name and the names of the assigned resources. If the task is in a synchronized SharePoint task list, **getTaskAsync** gets the task ID in the SharePoint list; otherwise, the SharePoint task ID is 0.</span></span>
    
     > [!NOTE]
     > <span data-ttu-id="e5014-p119">为了方便本文演示，示例代码有 bug。如果 **taskGuid** 未定义，**getTaskAsync** 函数不会显示错误。如果获得有效的任务 GUID，并接着选择其他任务，**getTaskAsync** 函数会获取 **getSelectedTaskAsync** 函数最近一次处理的任务的数据。</span><span class="sxs-lookup"><span data-stu-id="e5014-p119">For demonstration purposes, the example code includes a bug. If  **taskGuid** is undefined, the **getTaskAsync** function errors off. If you get a valid task GUID and then select a different task, the **getTaskAsync** function gets data for the most recent task that was operated on by the **getSelectedTaskAsync** function.</span></span>
  
   - <span data-ttu-id="e5014-p120">**getTaskFields**、**getResourceFields** 和 **getProjectFields** 是局部函数，通过多次调用 **getTaskFieldAsync**、**getResourceFieldAsync** 或 **getProjectFieldAsync**，以获取任务或资源的指定字段。在 project-15.debug.js 文件中，**ProjectTaskFields** 枚举和 **ProjectResourceFields** 枚举显示哪些字段受支持。</span><span class="sxs-lookup"><span data-stu-id="e5014-p120">**getTaskFields**, **getResourceFields**, and **getProjectFields** are local functions that call **getTaskFieldAsync**, **getResourceFieldAsync**, or **getProjectFieldAsync** multiple times to get specified fields of a task or a resource. In the project-15.debug.js file, the **ProjectTaskFields** enumeration and the **ProjectResourceFields** enumeration show which fields are supported.</span></span>
    
   - <span data-ttu-id="e5014-159">**getSelectedViewAsync** 函数用于获取视图类型（如 project-15.debug.js 的 **ProjectViewTypes** 枚举所定义）和视图名称。</span><span class="sxs-lookup"><span data-stu-id="e5014-159">The **getSelectedViewAsync** function gets the type of view (defined in the **ProjectViewTypes** enumeration in project-15.debug.js) and the name of the view.</span></span>
    
   - <span data-ttu-id="e5014-p121">如果项目与 SharePoint 任务列表同步，则 **getWSSUrlAsync** 函数获取任务列表的 URL 和名称。如果项目不与 SharePoint 任务列表同步，则 **getWSSUrlAsync** 函数错误关闭。</span><span class="sxs-lookup"><span data-stu-id="e5014-p121">If the project is synchronized with a SharePoint tasks list, the  **getWSSUrlAsync** function gets the URL and the name of the tasks list. If the project is not synchronized with a SharePoint tasks list, the **getWSSUrlAsync** function errors off.</span></span>
    
     > [!NOTE]
     > <span data-ttu-id="e5014-162">若要获取任务列表的 SharePoint URL 和名称，建议将 **getProjectFieldAsync** 函数与 [ProjectProjectFields](https://docs.microsoft.com/javascript/api/office/office.projectprojectfields) 枚举中的 **WSSUrl** 和 **WSSList** 常量配合使用。</span><span class="sxs-lookup"><span data-stu-id="e5014-162">To get the SharePoint URL and name of the tasks list, we recommend that you use the  **getProjectFieldAsync** function with the **WSSUrl** and **WSSList** constants in the [ProjectProjectFields](https://docs.microsoft.com/javascript/api/office/office.projectprojectfields) enumeration.</span></span>

   <span data-ttu-id="e5014-p122">以下代码的每个函数中都包含由 `function (asyncResult)` 指定的匿名函数，该函数是获取异步结果的回叫。你可以使用命名函数，而不是匿名函数，前者有助于实现复杂外接程序的可维护性。</span><span class="sxs-lookup"><span data-stu-id="e5014-p122">Each of the functions in the following code includes an anonymous function that is specified by  `function (asyncResult)`, which is a callback that gets the asynchronous result. Instead of anonymous functions, you could use named functions, which can help with maintainability of complex add-ins.</span></span>

    ```javascript
    // Get the data in the selected cells of the grid in the active view.
    function getSelectedDataAsync() {
        _projDoc.getSelectedDataAsync(
            Office.CoercionType.Text,
            { ValueFormat: "Formatted" },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded)
                    text.value = asyncResult.value;
                else
                    logMethodError("getSelectedDataAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        );
    }

    // Get the GUID of the selected task.
    function getSelectedTaskAsync() {
        _projDoc.getSelectedTaskAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                taskGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get the GUID of the selected resource.
    function getSelectedResourceAsync() {
        _projDoc.getSelectedResourceAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                resourceGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get data for the specified task.
    function getTaskAsync() {
        if (taskGuid != undefined) {
            _projDoc.getTaskAsync(
                taskGuid,
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        logMethodError("getTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                    } else {
                        var taskInfo = asyncResult.value;
                        var taskOutput = "Task name: " + taskInfo.taskName +
                                         "\nGUID: " + taskGuid +
                                         "\nWSS Id: " + taskInfo.wssTaskId +
                                         "\nResourceNames: " + taskInfo.resourceNames;
                        text.value = taskOutput;
                    }
                }
            );
        } else {
            text.value = 'Task GUID not valid:\n' + taskGuid;
        } 
    }

    // Get additional data for task fields.
    function getTaskFields() {
        text.value = "";

        _projDoc. getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Name: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.ID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "ID: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Start: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Duration,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Duration: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Priority,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Priority: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Notes,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Notes: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        ); 
    }

    // Get data for the specified resource fields.
    function getResourceFields() {
        text.value = "";

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Resource name: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Cost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.StandardRate,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Standard Rate: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualCost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualWork,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Work: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Units,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Units: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );
    }

    // Get the URL and list name of the synchronized SharePoint task list.
    // Recommended: use getProjectField instead.
    function getWSSUrlAsync() {
        _projDoc.getWSSUrlAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = "SharePoint URL:\n" + asyncResult.value.serverUrl
                    + "\nList name: " + asyncResult.value.listName;
            }
            else {
                logMethodError("getWSSUrlAsync", asyncResult.error.name, asyncResult.error.message);
            }
        });
    }

    // Get the type and name of the selected view.
    function getSelectedViewAsync() {
        _projDoc.getSelectedViewAsync(function (asyncResult) {
            text.value = "View type: " + asyncResult.value.viewType
                + "\nName: " + asyncResult.value.viewName;
        });
    }

    // Get information about the active project.
    function getProjectFields() {
        text.value = "";

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Project GUID: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nStart: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Finish,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nFinish: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProject " + errorText);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencyDigits,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nCurrency digits: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbol,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Currency symbol: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbolPosition,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSymbol position: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nProject web app URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSList,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint list: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    ```

4. <span data-ttu-id="e5014-p123">添加 JavaScript 事件处理程序回调和函数，以注册和取消注册任务选择、资源选择、视图选择更改事件处理程序。**manageEventHandlerAsync** 函数用于添加或删除指定的事件处理程序，具体视 _operation_ 参数而定。operation 的可取值为 **addHandlerAsync** 或 **removeHandlerAsync**。</span><span class="sxs-lookup"><span data-stu-id="e5014-p123">Add JavaScript event handler callbacks and functions to register the task selection, resource selection, and view selection change event handlers and to unregister the event handlers. The **manageEventHandlerAsync** function adds or removes the specified event handler, depending on the _operation_ parameter. The operation can be **addHandlerAsync** or **removeHandlerAsync**.</span></span>
    
   <span data-ttu-id="e5014-168">**manageTaskEventHandler**、**manageResourceEventHandler** 和 **manageViewEventHandler** 函数可以添加或删除 _docMethod_ 参数指定的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e5014-168">The **manageTaskEventHandler**, **manageResourceEventHandler**, and **manageViewEventHandler** functions can add or remove an event handler, as specified by the _docMethod_ parameter.</span></span>

    ```javascript
    // Task selection changed event handler.
    function onTaskSelectionChanged(eventArgs) {
        text.value = "In task selection change event handler";
    }

    // Resource selection changed event handler.
    function onResourceSelectionChanged(eventArgs) {
        text.value = "In Resource selection changed event handler";
    }

    // View selection changed event handler.
    function onViewSelectionChanged(eventArgs) {
        text.value = "In View selection changed event handler";
    }

    // Add or remove the specified event handler.
    function manageEventHandlerAsync(eventType, handler, operation, onComplete) {
        _projDoc[operation]   //The operation is addHandlerAsync or removeHandlerAsync.
        (
            eventType,
            handler,
            function (asyncResult) {
                if (onComplete) {
                    onComplete(asyncResult, operation);
                } else {
                    var message = "Operation: " + operation;
                    message = message + "\nStatus: " + asyncResult.status + "\n";
                    text.value = message;
                }
            }
        );
    }

    // Write the asyncResult status from the manageEventHandlerAsync function (optional). 
    function onComplete(asyncResult, operation) {
        var message = "In onComplete function for " + operation;
        message = message + "\nStatus: " + asyncResult.status;
        text.value = message;
    }

    // Add or remove a task selection changed event handler.
    function manageTaskEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.TaskSelectionChanged,      // The task selection changed event.
            onTaskSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a resource selection changed event handler.
    function manageResourceEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ResourceSelectionChanged,  // The resource selection changed event.
            onResourceSelectionChanged,                 // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a view selection changed event handler.
    function manageViewEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ViewSelectionChanged,      // The view selection changed event.
            onViewSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }
    ```

5. <span data-ttu-id="e5014-p124">对于 HTML 文档正文，添加调用 JavaScript 函数的按钮进行测试。例如，在公共 JSOM API 的 **div** 元素中，添加调用普通 **getSelectedDataAsync** 函数的输入按钮。</span><span class="sxs-lookup"><span data-stu-id="e5014-p124">For the body of the HTML document, add buttons that call the JavaScript functions for testing. For example, in the  **div** element for the common JSOM API, add an input button that calls the general **getSelectedDataAsync** function.</span></span>
    
    ```HTML
    <body>
        <div id="Common_JSOM_API">
        OBJECT MODEL TESTS
        <br /><br />       
        <strong>General function:</strong>
        <br />
        <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
            value="getSelectedDataAsync" />
        </div>
        <!--  more code . . .  -->
    ```

6. <span data-ttu-id="e5014-171">添加 **div** 部分，其中包含项目专用任务函数和 **TaskSelectionChanged** 事件的按钮。</span><span class="sxs-lookup"><span data-stu-id="e5014-171">Add a **div** section with buttons for project-specific task functions and for the **TaskSelectionChanged** event.</span></span>
    
    ```HTML
    <div id="ProjectSpecificTask">
      <br />
      <strong>Project-specific task methods:</strong><br />
      <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
      <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
      <strong>Task selection changed:</strong>
      <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>         
    </div>
    ```

7. <span data-ttu-id="e5014-172">添加 **div** 部分，其中包含资源方法和事件、视图方法和事件、项目属性和上下文属性的按钮</span><span class="sxs-lookup"><span data-stu-id="e5014-172">Add  **div** sections with buttons for the resource methods and events, view methods and events, project properties, and context properties</span></span>
    
    ```HTML
    <div id="ResourceMethods">
      <br />
      <strong>Resource methods:</strong>
      <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
      <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
      <strong>Resource selection changed:</strong>
      <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    <div id="ViewMethods">
      <br />
      <strong>View method:</strong>
      <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
      <strong>View selection changed:</strong>
      <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>         
    </div>
    <div id="ProjectMethods">
      <br />
      <strong>Project properties:</strong>
      <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
    </div>
    <div id="ContextVariables">
      <br />
      <strong>Context properties:</strong>
      <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
    </div>
    ```

8. <span data-ttu-id="e5014-p125">要设置按钮元素的格式，可添加 CSS **style** 元素。例如，添加以下内容作为 **head** 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="e5014-p125">To format the button elements, add a CSS  **style** element. For example, add the following as a child of the **head** element.</span></span>
    
    ```HTML
    <style type="text/css">
        .button-wide
        {
            width: 210px;
            margin-top: 2px;
        }
        .button-narrow
        {
            width: 80px;
            margin-top: 2px;
        }
    </style>
    ```

<span data-ttu-id="e5014-175">过程 3 显示如何安装和使用 Project OM Test 加载项功能。</span><span class="sxs-lookup"><span data-stu-id="e5014-175">Procedure 3 shows how to install and use the Project OM Test add-in features.</span></span>

## <a name="procedure-3-to-install-and-use-the-project-om-test-add-in"></a><span data-ttu-id="e5014-p126">过程 3. 安装和使用 Project OM Test 加载项</span><span class="sxs-lookup"><span data-stu-id="e5014-p126">Procedure 3. To install and use the Project OM Test add-in</span></span>

1. <span data-ttu-id="e5014-p127">为包含 JSOM_SimpleOMCalls.XML 清单的目录创建一个文件共享。可以在本地计算机或可通过网络访问的远程计算机上创建该文件共享。例如，如果清单位于本地计算机上的 `C:\Project\AppManifests` 目录中，则运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="e5014-p127">Create a file share for the directory that contains the JSOM_SimpleOMCalls.xml manifest. You can create the file share on the local computer or on a remote computer that is accessible on the network. For example, if the manifest is in the  `C:\Project\AppManifests` directory on the local computer, run the following command:</span></span>
    
    `Net share AppManifests=C:\Project\AppManifests`
    
2. <span data-ttu-id="e5014-p128">为包含 Project OM Test 加载项的 HTML 和 JavaScript 文件的目录创建一个文件共享。确保文件共享路径与在 JSOM_SimpleOMCalls.xml 清单中指定的路径匹配。例如，如果文件位于本地计算机上的 `C:\Project\AppSource` 目录中，则运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="e5014-p128">Create a file share for the directory that contains the HTML and JavaScript files for the Project OM Test add-in. Ensure the file share path matches the path that is specified in the JSOM_SimpleOMCalls.xml manifest. For example, if the files are in the  `C:\Project\AppSource` directory on the local computer, run the following command:</span></span>
    
    `net share AppSource=C:\Project\AppSource`

3. <span data-ttu-id="e5014-184">在 Project 中，打开“Project 选项”\*\*\*\* 对话框，再依次选择“信任中心”\*\*\*\* 和“信任中心设置”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e5014-184">In Project, open the **Project Options** dialog box, choose **Trust Center**, and then choose  **Trust Center Settings**.</span></span>
    
   <span data-ttu-id="e5014-185">[Project 任务窗格加载项](../project/project-add-ins.md)中还介绍了加载项注册过程和其他详细信息。</span><span class="sxs-lookup"><span data-stu-id="e5014-185">The procedure for registering an add-in is also described in [Task pane add-ins for Project](../project/project-add-ins.md), with additional information.</span></span>
    
4. <span data-ttu-id="e5014-186">在“信任中心”\*\*\*\* 对话框的左侧窗格中，选择“受信任的加载项目录”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e5014-186">In the **Trust Center** dialog box, in the left pane, choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="e5014-p129">如果已添加必应搜索加载项的 `\\ServerName\AppManifests` 路径，请跳过这一步。否则，在“受信任的加载项目录”\*\*\*\* 窗格中，向“目录 URL”\*\*\*\* 文本框添加 `\\ServerName\AppManifests` 路径，选择“添加目录”\*\*\*\*，将网络共享启用为默认源（见图 1），再选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="e5014-p129">If you have already added the `\\ServerName\AppManifests` path for the Bing Search add-in, skip this step. Otherwise, in the **Trusted Add-in Catalogs** pane, add the `\\ServerName\AppManifests` path in the **Catalog Url** text box, choose **Add catalog**, enable the network share as a default source (see Figure 1), and then choose  **OK**.</span></span>
    
   <span data-ttu-id="e5014-189">*图 1：添加加载项清单的网络文件共享*</span><span class="sxs-lookup"><span data-stu-id="e5014-189">*Figure 1. Adding a network file share for add-in manifests*</span></span>

   ![为应用程序清单添加网络文件共享](../images/pj15-create-simple-agave-manage-catalogs.png)

6. <span data-ttu-id="e5014-p130">添加新的外接程序或更改源代码后，重新启动项目。在“**项目**”功能区，选择“**Office 外接程序**”下拉菜单，然后选择“**查看所有**”。在“**插入外接程序**”对话框中，选择“**共享文件夹**”（见图 2），选择“**Project OM Test**”，然后选择“**插入**”。Project OM Test 外接程序在任务窗格启动。</span><span class="sxs-lookup"><span data-stu-id="e5014-p130">After you add new add-ins, or change the source code, restart Project. On the  **PROJECT** ribbon, choose the **Office Add-ins** drop-down menu, and then choose **See All**. In the  **Insert Add-in** dialog box, choose **SHARED FOLDER** (see Figure 2), select **Project OM Test**, and then choose  **Insert**. The Project OM Test add-in starts in a task pane.</span></span>
    
   <span data-ttu-id="e5014-195">*图 2：启动文件共享上的“Project OM 测试”加载项*</span><span class="sxs-lookup"><span data-stu-id="e5014-195">*Figure 2. Starting the Project OM Test add-in that is on a file share*</span></span>

   ![插入应用程序](../images/pj15-create-simple-agave-start-agave-app.png)

7. <span data-ttu-id="e5014-p131">在 Project 中，创建并保存具有至少两个任务的简单项目。例如，创建名为 T1、T2 的任务和名为 M1 的里程碑，然后将任务工期和前置任务设置为与图 3 中的类似。选择功能区上的“**项目**”选项卡，选择任务 T2 的整个行，然后在任务窗格中选择“**getSelectedDataAsync**”按钮。图 3 显示在 **Project OM Test** 外接程序的文本框中选择的数据。</span><span class="sxs-lookup"><span data-stu-id="e5014-p131">In Project, create and save a simple project that has at least two tasks. For example, create tasks named T1, T2, and a milestone named M1, and then set the task durations and predecessors to be similar to those in Figure 3. Choose the  **PROJECT** tab on the ribbon, select the entire row for task T2, and then choose the **getSelectedDataAsync** button in the task pane. Figure 3 shows the data that is selected in the text box of the **Project OM Test** add-in.</span></span>
    
   <span data-ttu-id="e5014-201">*图 3：使用“Project OM 测试”加载项*</span><span class="sxs-lookup"><span data-stu-id="e5014-201">*Figure 3. Using the Project OM Test add-in*</span></span>

   ![使用 Project OM Test 应用程序](../images/pj15-create-simple-agave-project-om-test.png)

8. <span data-ttu-id="e5014-p132">选择第一项任务的“**工期**”列中的单元格，然后选择 **Project OM Test** 外接程序中的“**getSelectedDataAsync**”按钮。**getSelectedDataAsync** 函数将文本框值设置为显示 `2 days`。</span><span class="sxs-lookup"><span data-stu-id="e5014-p132">Select the cell in the  **Duration** column for the first task, and then choose the **getSelectedDataAsync** button in the **Project OM Test** add-in. The **getSelectedDataAsync** function sets the text box value to show `2 days`.</span></span> 
    
9. <span data-ttu-id="e5014-p133">选择所有三项任务的三个**工期**单元格。**getSelectedDataAsync** 函数为在不同行中选定的单元格返回以分号分隔的文本值，例如，`2 days;4 days;0 days`。</span><span class="sxs-lookup"><span data-stu-id="e5014-p133">Select the three  **Duration** cells for all three tasks. The **getSelectedDataAsync** function returns semicolon-separated text values for cells selected in different rows, for example, `2 days;4 days;0 days`.</span></span>
    
   <span data-ttu-id="e5014-p134">**getSelectedDataAsync** 函数返回行中选定单元格的以逗号分隔的文本值。有关图 3 中的示例，选中任务 T2 的整行。在选择 **getSelectedDataAsync** 时，文本框显示以下内容：`,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`</span><span class="sxs-lookup"><span data-stu-id="e5014-p134">The  **getSelectedDataAsync** function returns comma-separated text values for cells selected within a row. For example in Figure 3, the entire row for task T2 is selected. When you choose **getSelectedDataAsync**, the text box shows the following:  `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`</span></span>
    
   <span data-ttu-id="e5014-p135">“**标记**”列和“**资源名称**”列均为空，因此，文本数组显示这些列的空值。`<NA>` 值代表“**添加新列**”单元格。</span><span class="sxs-lookup"><span data-stu-id="e5014-p135">The  **Indicators** column and the **Resource Names** column are both empty, so the text array shows empty values for those columns. The `<NA>` value is for the **Add New Column** cell.</span></span>
    
10. <span data-ttu-id="e5014-p136">选择任务 T2 行中的任何单元格，或任务 T2 的整行，然后选择 **getSelectedTaskAsync**。文本框显示任务 GUID 值，例如，`{25D3E03B-9A7D-E111-92FC-00155D3BA208}`。Project 在 **Project OM Test** 加载项的全局 **taskGuid** 变量中存储该值。</span><span class="sxs-lookup"><span data-stu-id="e5014-p136">Select any cell in the row for task T2, or the entire row for task T2, and then choose  **getSelectedTaskAsync**. The text box shows the task GUID value, for example,  `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`. Project stores that value in the global  **taskGuid** variable of the **Project OM Test** add-in.</span></span>
    
11. <span data-ttu-id="e5014-p137">选择“getTaskAsync”\*\*\*\*。如果 **taskGuid** 变量包含任务 T2 的 GUID，文本框中会显示任务信息。**ResourceNames** 值为空。</span><span class="sxs-lookup"><span data-stu-id="e5014-p137">Select **getTaskAsync**. If the **taskGuid** variable contains the GUID for task T2, the text box displays the task information. The **ResourceNames** value is empty.</span></span>
    
    <span data-ttu-id="e5014-p138">创建两个本地资源 R1 和 R2，将其分配给任务 T2（每个分配 50%），然后重新选择 **getTaskAsync**。文本框中的结果包含资源信息。如果任务位于同步的 SharePoint 任务列表中，那么结果还会包含 SharePoint 任务 ID。</span><span class="sxs-lookup"><span data-stu-id="e5014-p138">Create two local resources R1 andR2, assign them to task T2 at 50% each, and choose  **getTaskAsync** again. The results in the text box include the resource information. If the task is in a synchronized SharePoint task list, the results also include the SharePoint task ID.</span></span>
    
    - <span data-ttu-id="e5014-221">任务名称：`T2`</span><span class="sxs-lookup"><span data-stu-id="e5014-221">Task name: `T2`</span></span>
    - <span data-ttu-id="e5014-222">GUID：`{25D3E03B-9A7D-E111-92FC-00155D3BA208}`</span><span class="sxs-lookup"><span data-stu-id="e5014-222">GUID: `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`</span></span>
    - <span data-ttu-id="e5014-223">WSS ID：`0`</span><span class="sxs-lookup"><span data-stu-id="e5014-223">WSS Id: `0`</span></span>
    - <span data-ttu-id="e5014-224">ResourceNames: `R1[50%],R2[50%]`</span><span class="sxs-lookup"><span data-stu-id="e5014-224">ResourceNames: `R1[50%],R2[50%]`</span></span>

12. <span data-ttu-id="e5014-p139">选择“获取任务字段”\*\*\*\* 按钮。**getTaskFields** 函数会多次调用 **getTaskfieldAsync** 函数，以获取任务名称、索引、开始日期、持续时间、优先级和任务备注。</span><span class="sxs-lookup"><span data-stu-id="e5014-p139">Select the **Get Task Fields** button. The **getTaskFields** function calls the **getTaskfieldAsync** function multiple times for the task name, index, start date, duration, priority, and task notes.</span></span>

    - <span data-ttu-id="e5014-227">名称：`T2`</span><span class="sxs-lookup"><span data-stu-id="e5014-227">Name: `T2`</span></span>
    - <span data-ttu-id="e5014-228">ID：`2`</span><span class="sxs-lookup"><span data-stu-id="e5014-228">ID: `2`</span></span>
    - <span data-ttu-id="e5014-229">开始日期：`Thu 6/14/12`</span><span class="sxs-lookup"><span data-stu-id="e5014-229">Start: `Thu 6/14/12`</span></span>
    - <span data-ttu-id="e5014-230">持续时间：`4d`</span><span class="sxs-lookup"><span data-stu-id="e5014-230">Duration: `4d`</span></span>
    - <span data-ttu-id="e5014-231">优先级：`500`</span><span class="sxs-lookup"><span data-stu-id="e5014-231">Priority: `500`</span></span>
    - <span data-ttu-id="e5014-232">备注：此为任务 T2 的备注。</span><span class="sxs-lookup"><span data-stu-id="e5014-232">Notes: This is a note for task T2.</span></span> <span data-ttu-id="e5014-233">仅为测试备注。</span><span class="sxs-lookup"><span data-stu-id="e5014-233">It is only a test note.</span></span> <span data-ttu-id="e5014-234">若为实际备注，应包含一些真实信息。</span><span class="sxs-lookup"><span data-stu-id="e5014-234">If it had been a real note, there would be some real information.</span></span>

13. <span data-ttu-id="e5014-p141">选择“getWSSUrlAsync”\*\*\*\* 按钮。如果项目是以下类型之一，结果中显示任务列表 URL 和名称。</span><span class="sxs-lookup"><span data-stu-id="e5014-p141">Select the **getWSSUrlAsync** button. If the project is one of the following kinds, the results show the task list URL and name.</span></span>
    
    - <span data-ttu-id="e5014-237">导入到 Project Server 的 SharePoint 任务列表。</span><span class="sxs-lookup"><span data-stu-id="e5014-237">A SharePoint task list that was imported to Project Server.</span></span>
    - <span data-ttu-id="e5014-238">导入到 Project Professional，再保存回 SharePoint（未使用 Project Server）的 SharePoint 任务列表。</span><span class="sxs-lookup"><span data-stu-id="e5014-238">A SharePoint task list that was imported to Project Professional, and then saved back in SharePoint (not using Project Server).</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="e5014-239">如果 Project Professional 安装在 Windows Server 计算机上，若要将项目保存回 SharePoint，可使用“服务器管理器”\*\*\*\* 添加“桌面体验”\*\*\*\* 功能。</span><span class="sxs-lookup"><span data-stu-id="e5014-239">If Project Professional is installed on a Windows Server computer, to be able to save the project back to SharePoint, you can use the  **Server Manager** to add the **Desktop Experience** feature.</span></span>

    <span data-ttu-id="e5014-240">如果项目是本地项目，或者如果你使用 Project Professional 打开由 Project Server 管理的项目，那么 **getWSSUrlAsync** 方法会显示一个未定义错误。</span><span class="sxs-lookup"><span data-stu-id="e5014-240">If the project is a local project, or if you use Project Professional to open a project that is managed by Project Server, the  **getWSSUrlAsync** method shows an undefined error.</span></span>

    - <span data-ttu-id="e5014-241">SharePoint URL：`http://ServerName`</span><span class="sxs-lookup"><span data-stu-id="e5014-241">SharePoint URL: `http://ServerName`</span></span>
    - <span data-ttu-id="e5014-242">列表名称：`Test task list`</span><span class="sxs-lookup"><span data-stu-id="e5014-242">List name: `Test task list`</span></span>
    

14. <span data-ttu-id="e5014-p142">选择“TaskSelectionChanged 事件”\*\*\*\* 部分中的“添加”\*\*\*\* 按钮，这会调用 **manageTaskEventHandler** 函数，以注册任务选择更改事件，并在文本框中返回 `In onComplete function for addHandlerAsync Status: succeeded`。选择其他任务；此时，文本框会显示 `In task selection changed event handler`，这是任务选择更改事件的回调函数输出。选择“删除”\*\*\*\* 按钮可以取消注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e5014-p142">Select the **Add** button in the **TaskSelectionChanged event** section, which calls the **manageTaskEventHandler** function to register a task selection changed event and returns `In onComplete function for addHandlerAsync Status: succeeded` in the text box. Select a different task; the text box shows `In task selection changed event handler`, which is the output of the callback function for the task selection changed event. Choose the  **Remove** button to unregister the event handler.</span></span>
    
15. <span data-ttu-id="e5014-p143">若要使用资源方法，首先选择视图（如“**资源工作表**”、“**资源使用状况**”或“**资源窗体**”），然后选择该视图中的资源。选择 **getSelectedResourceAsync** 以初始化 **resourceGuid** 变量，然后选择“**获取资源域**”以对资源属性的 **getResourceFieldAsync** 进行多次调用。还可以添加或删除资源选择更改事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e5014-p143">To use the resource methods, first select a view such as  **Resource Sheet**,  **Resource Usage**, or  **Resource Form**, and then select a resource in that view. Choose  **getSelectedResourceAsync** to initialize the **resourceGuid** variable, and then choose **Get Resource Fields** to call **getResourceFieldAsync** multiple times for the resource properties. You can also add or remove the resource selection changed event handler.</span></span>
    
    - <span data-ttu-id="e5014-249">资源名称：`R1`</span><span class="sxs-lookup"><span data-stu-id="e5014-249">Resource name: `R1`</span></span>
    - <span data-ttu-id="e5014-250">成本：`$800.00`</span><span class="sxs-lookup"><span data-stu-id="e5014-250">Cost: `$800.00`</span></span>
    - <span data-ttu-id="e5014-251">标准费率：`$50.00/h`</span><span class="sxs-lookup"><span data-stu-id="e5014-251">Standard Rate: `$50.00/h`</span></span>
    - <span data-ttu-id="e5014-252">实际成本：`$0.00`</span><span class="sxs-lookup"><span data-stu-id="e5014-252">Actual Cost: `$0.00`</span></span>
    - <span data-ttu-id="e5014-253">实际工时：`0h`</span><span class="sxs-lookup"><span data-stu-id="e5014-253">Actual Work: `0h`</span></span>
    - <span data-ttu-id="e5014-254">单位：`100%`</span><span class="sxs-lookup"><span data-stu-id="e5014-254">Units: `100%`</span></span>

16. <span data-ttu-id="e5014-p144">选择“getSelectedViewAsync”\*\*\*\*，显示活动视图的类型和名称。还可以添加或删除视图选择更改事件处理程序。例如，如果“资源表单”\*\*\*\* 是活动视图，**getSelectedViewAsync** 函数会在文本框中显示以下内容：</span><span class="sxs-lookup"><span data-stu-id="e5014-p144">Select **getSelectedViewAsync** to show the type and name of the active view. You can also add or remove the view selection changed event handler. For example, if **Resource Form** is the active view, the **getSelectedViewAsync** function shows the following in the text box:</span></span>
    
    - <span data-ttu-id="e5014-258">视图类型：`6`</span><span class="sxs-lookup"><span data-stu-id="e5014-258">View type: `6`</span></span>
    - <span data-ttu-id="e5014-259">名称：`Resource Form`</span><span class="sxs-lookup"><span data-stu-id="e5014-259">Name: `Resource Form`</span></span>
    
17. <span data-ttu-id="e5014-p145">选择“获取项目字段”\*\*\*\*，以多次调用 **getProjectFieldAsync** 函数来获取有效项目的不同属性。如果项目是从 Project Web App 打开，**getProjectFieldAsync** 函数可以获取 Project Web App 实例的 URL。</span><span class="sxs-lookup"><span data-stu-id="e5014-p145">Select **Get Project Fields** to call the **getProjectFieldAsync** function multiple times for different properties of the active project. If the project is opened from Project Web App, the **getProjectFieldAsync** function can get the URL of the Project Web App instance.</span></span>
    
    - <span data-ttu-id="e5014-262">项目 GUID：`9845922E-DAB4-E111-8AF3-00155D3BA208`</span><span class="sxs-lookup"><span data-stu-id="e5014-262">Project GUID: `9845922E-DAB4-E111-8AF3-00155D3BA208`</span></span>
    - <span data-ttu-id="e5014-263">开始日期：`Tue 6/12/12`</span><span class="sxs-lookup"><span data-stu-id="e5014-263">Start: `Tue 6/12/12`</span></span>
    - <span data-ttu-id="e5014-264">完成日期：`Tue 6/19/12`</span><span class="sxs-lookup"><span data-stu-id="e5014-264">Finish: `Tue 6/19/12`</span></span>
    - <span data-ttu-id="e5014-265">货币位数：`2`</span><span class="sxs-lookup"><span data-stu-id="e5014-265">Currency digits: `2`</span></span>
    - <span data-ttu-id="e5014-266">货币符号：`$`</span><span class="sxs-lookup"><span data-stu-id="e5014-266">Currency symbol: `$`</span></span>
    - <span data-ttu-id="e5014-267">符号位置：`0`</span><span class="sxs-lookup"><span data-stu-id="e5014-267">Symbol position: `0`</span></span>
    - <span data-ttu-id="e5014-268">Project Web App URL：`http://servername/pwa`</span><span class="sxs-lookup"><span data-stu-id="e5014-268">Project web app URL: `http://servername/pwa`</span></span>
  
18. <span data-ttu-id="e5014-p146">选择“获取上下文值”\*\*\*\* 按钮，获取 **Office.Context.document** 对象和 **Office.context.application** 对象的属性，从而获取运行加载项的文档和应用的属性。例如，如果 Project1.mpp 文件位于本地计算机桌面上，文档 URL 为 `C:\Users\UserAlias\Desktop\Project1.mpp`。如果 .mpp 文件位于 SharePoint 库中，值为文档的 URL。如果使用 Project Professional 2013 从 Project Web App 打开 Project1 项目，文档 URL 为 `<>\Project1`。</span><span class="sxs-lookup"><span data-stu-id="e5014-p146">Select  the **Get Context Values** button get properties of the document and the application in which the add-in is running, by getting properties of the **Office.Context.document** object and the **Office.context.application** object. For example, if the Project1.mpp file is on the local computer desktop, the document URL is `C:\Users\UserAlias\Desktop\Project1.mpp`. If the .mpp file is in a SharePoint library, the value is the URL of the document. If you use Project Professional 2013 to open a project named Project1 from Project Web App, the document URL is  `<>\Project1`.</span></span>
    
    - <span data-ttu-id="e5014-273">文档 URL：`<>\Project1`</span><span class="sxs-lookup"><span data-stu-id="e5014-273">Document URL: `<>\Project1`</span></span>
    - <span data-ttu-id="e5014-274">文档模式：`readWrite`</span><span class="sxs-lookup"><span data-stu-id="e5014-274">Document mode: `readWrite`</span></span>
    - <span data-ttu-id="e5014-275">应用语言：`en-US`</span><span class="sxs-lookup"><span data-stu-id="e5014-275">App language: `en-US`</span></span>
    - <span data-ttu-id="e5014-276">显示语言：`en-US`</span><span class="sxs-lookup"><span data-stu-id="e5014-276">Display language: `en-US`</span></span>
    
19. <span data-ttu-id="e5014-p147">可以通过关闭并重启 Project 以在编辑源代码后刷新外接程序。在“**项目**”功能区中，“**Office 外接程序**”下拉列表维护最近使用的外接程序列表。</span><span class="sxs-lookup"><span data-stu-id="e5014-p147">You can refresh the add-in after you edit the source code by closing and restarting Project. In the  **Project** ribbon, the **Office Add-ins** drop-down list maintains the list of recently used add-ins.</span></span>
    
## <a name="example"></a><span data-ttu-id="e5014-279">示例</span><span class="sxs-lookup"><span data-stu-id="e5014-279">Example</span></span>

<span data-ttu-id="e5014-p148">Project 2013 SDK 下载包含 JSOMCall.html 文件、JSOM_Sample.js 文件和相关 Office.js、Office.debug.js、Project-15.js、Project-15.debug.js 文件的完整代码。以下是 JSOMCall.html 文件中的代码。</span><span class="sxs-lookup"><span data-stu-id="e5014-p148">The Project 2013 SDK download contains the complete code in the JSOMCall.html file, the JSOM_Sample.js file, and the related Office.js, Office.debug.js, Project-15.js, and Project-15.debug.js files. Following is the code in the JSOMCall.html file.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <title>Project OM Sample Code</title>
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

        <script type="text/javascript" src="MicrosoftAjax.js"></script>

        <!-- Use the CDN reference to office.js when deploying your add-in. -->
        <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script> -->
        <script type="text/javascript" src="Office.js"></script>
        <script type="text/javascript" src="JSOM_Sample.js"></script>

        <style type="text/css">           
            .button-wide {
                width: 210px;
                margin-top: 2px;
            }
            .button-narrow 
            {
                width: 80px;
                margin-top: 2px;
            }
        </style>
    </head>

    <body>
        <div id="Common_JSOM_API">
            OBJECT MODEL TESTS
            <br /><br />       
            <strong>General method:</strong>
            <br />
            <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
                value="getSelectedDataAsync" />
        </div>
        <div id="ProjectSpecificTask">
            <br />
            <strong>Project-specific task methods:</strong><br />
            <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
            <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
            <strong>Task selection changed:</strong>
            <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>         
        </div>
        <div id="ResourceMethods">
            <br />
            <strong>Resource methods:</strong>
            <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
            <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
            <strong>Resource selection changed:</strong>
            <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ViewMethods">
            <br />
            <strong>View method:</strong>
            <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
            <strong>View selection changed:</strong>
            <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>         
        </div>
        <div id="ProjectMethods">
            <br />
            <strong>Project properties:</strong>
            <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
        </div>
        <div id="ContextVariables">
            <br />
            <strong>Context properties:</strong>
            <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
        </div>
        <br />
        <textarea id="text" rows="10" cols="25">This is the text result.</textarea>
    </body>
</html
```

## <a name="robust-programming"></a><span data-ttu-id="e5014-282">可靠编程</span><span class="sxs-lookup"><span data-stu-id="e5014-282">Robust programming</span></span>

<span data-ttu-id="e5014-p149">**Project OM Test** 加载项是一个示例，显示如何使用 Project-15.js 和 Office.js 文件中 Project 2013 的某些 JavaScript 函数。此示例仅供测试用，不包括可靠的错误检查。例如，如果你未选择资源而运行 **getSelectedResourceAsync** 函数，则 **resourceGuid** 变量不进行初始化，并且对 **getResourceFieldAsync** 的调用将返回错误。对于生产加载项，应检查特定错误并忽略结果，隐藏未应用的功能，或通知用户选择视图并在使用函数前先进行有效选择。</span><span class="sxs-lookup"><span data-stu-id="e5014-p149">The  **Project OM Test** add-in is an example that shows the use of some JavaScript functions for Project 2013 in the Project-15.js and Office.js files. The example is for testing only and does not include robust error checks. For example, if you do not select a resource and run the **getSelectedResourceAsync** function, the **resourceGuid** variable is not initialized, and calls to **getResourceFieldAsync** return an error. For a production add-in, you should check for specific errors and ignore the results, hide functionality that does not apply, or notify the user to choose a view and make a valid selection before using a function.</span></span>

<span data-ttu-id="e5014-287">对于简单示例，下列代码中的错误输出包括 **actionMessage** 变量，该变量指定为避免 **getSelectedResourceAsync** 函数出错而执行的操作。</span><span class="sxs-lookup"><span data-stu-id="e5014-287">For a simple example, the error output in the following code includes the  **actionMessage** variable that specifies the action to take to avoid an error in the **getSelectedResourceAsync** function.</span></span>

```javascript
function logError(errorText) {
    text.value = "Error in " + errorText;
}

function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);
}

// Get the GUID of the selected resource.
function getSelectedResourceAsync() {
    _projDoc.getSelectedResourceAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = asyncResult.value;
            resourceGuid = asyncResult.value;
        }
        else {
            var actionMessage = "Select a resource before running the getSelectedResourceAsync method.";
            logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                               asyncResult.error.message, actionMessage);
        }
    });
}
```

<span data-ttu-id="e5014-288">Project 2013 SDK 下载中的 **HelloProject_OData** 示例包含使用 JQuery 库来显示弹出错误消息的 SurfaceErrors.js 文件。</span><span class="sxs-lookup"><span data-stu-id="e5014-288">The **HelloProject_OData** sample in the Project 2013 SDK download includes the SurfaceErrors.js file that uses the JQuery library to display a pop-up error message.</span></span> <span data-ttu-id="e5014-289">图 4 显示“toast”通知中的错误消息。</span><span class="sxs-lookup"><span data-stu-id="e5014-289">Figure 4 shows the error message in a "toast" notification.</span></span>

<span data-ttu-id="e5014-290">SurfaceErrors.js 文件中的以下代码包括创建 **Toast** 对象的 **throwError** 函数。</span><span class="sxs-lookup"><span data-stu-id="e5014-290">The following code in the SurfaceErrors.js file includes the  **throwError** function that creates a **Toast** object.</span></span>

```javascript
/*
 * Show error messages in a "toast" notification.
 */

// Throws a custom defined error.
function throwError(errTitle, errMessage) {
    try {
        // Define and throw a custom error.
        var customError = { name: errTitle, message: errMessage }
        throw customError;
    }
    catch (err) {
        // Catch the error and display it to the user.
        Toast.showToast(err.name, err.message);
    }
}

// Add a dynamically-created div "toast" for displaying errors to the user.
var Toast = {

    Toast: "divToast",
    Close: "btnClose",
    Notice: "lblNotice",
    Output: "lblOutput",

    // Show the toast with the specified information.
    showToast: function (title, message) {

        if (document.getElementById(this.Toast) == null) {
            this.createToast();
        }

        document.getElementById(this.Notice).innerText = title;
        document.getElementById(this.Output).innerText = message;

        $("#" + this.Toast).hide();
        $("#" + this.Toast).show("slow");
    },

    // Create the display for the toast.
    createToast: function () {
        var divToast;
        var lblClose;
        var btnClose;
        var divOutput;
        var lblOutput;
        var lblNotice;

        // Create the container div.
        divToast = document.createElement("div");
        var toastStyle = "background-color:rgba(220, 220, 128, 0.80);" +
            "position:absolute;" +
            "bottom:0px;" +
            "width:90%;" +
            "text-align:center;" +
            "font-size:11pt;";
        divToast.setAttribute("style", toastStyle);
        divToast.setAttribute("id", this.Toast);

        // Create the close button.
        lblClose = document.createElement("div");
        lblClose.setAttribute("id", this.Close);
        var btnStyle = "text-align:right;" +
            "padding-right:10px;" +
            "font-size:10pt;" +
            "cursor:default";
        lblClose.setAttribute("style", btnStyle);
        lblClose.appendChild(document.createTextNode("CLOSE "));

        btnClose = document.createElement("span");
        btnClose.setAttribute("style", "cursor:pointer;");
        btnClose.setAttribute("onclick", "Toast.close()");
        btnClose.innerText = "X";
        lblClose.appendChild(btnClose);

        // Create the div to contain the toast title and message.
        divOutput = document.createElement("div");
        divOutput.setAttribute("id", "divOutput");
        var outputStyle = "margin-top:0px;";
        divOutput.setAttribute("style", outputStyle);

        lblNotice = document.createElement("span");
        lblNotice.setAttribute("id", this.Notice);
        var labelStyle = "font-weight:bold;margin-top:0px;";
        lblNotice.setAttribute("style", labelStyle);

        lblOutput = document.createElement("span");
        lblOutput.setAttribute("id", this.Output);

        // Add the child nodes to the toast div.
        divOutput.appendChild(lblNotice);
        divOutput.appendChild(document.createElement("br"));
        divOutput.appendChild(lblOutput);
        divToast.appendChild(lblClose);
        divToast.appendChild(divOutput);

        // Add the toast div to the document body.
        document.body.appendChild(divToast);
    },

    // Close the toast.
    close: function () {
        $("#" + this.Toast).hide("slow");
    }
}
```

<span data-ttu-id="e5014-291">要使用 **throwError** 函数，可在 JSOMCall.html 文件中包括 JQuery 库和 SurfaceErrors.js 脚本，然后在其他 JavaScript 函数（如 **logMethodError**）中添加对 **throwError** 的调用。</span><span class="sxs-lookup"><span data-stu-id="e5014-291">To use the  **throwError** function, include the JQuery library and the SurfaceErrors.js script in the JSOMCall.html file, and then add a call to **throwError** in other JavaScript functions such as **logMethodError**.</span></span>

> [!NOTE]
> <span data-ttu-id="e5014-p151">部署加载项之前，请将 office.js 引用和 jQuery 引用更改为内容发布网络 (CDN) 引用。CDN 引用可提供最新的版本和更好的性能。</span><span class="sxs-lookup"><span data-stu-id="e5014-p151">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

```HTML
<!DOCTYPE html>
<html>
<head>
    <title>Project OM Sample Code</title>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <script type="text/javascript" src="MicrosoftAjax.js"></script>

    <!-- Use the CDN reference to Office.js and jQuery when deploying your add-in. -->
    <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script> -->
    <script type="text/javascript" src="Office.js"></script>
    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jQuery/jquery-1.9.0.min.js"></script>

    <script type="text/javascript" src="JSOM_Sample.js"></script>
    <script type="text/javascript" src="SurfaceErrors.js"></script>

    <!-- . . . INVALID USE OF SYMBOLS . . . -->
</head>

```

<br/>


```javascript
function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);

    throwError(methodName + " error", actionMessage);
}
```

<br/>

<span data-ttu-id="e5014-294">*图 4：SurfaceErrors.js 文件中的函数可以显示“toast”通知*</span><span class="sxs-lookup"><span data-stu-id="e5014-294">*Figure 4. Functions in the SurfaceErrors.js file can show a "toast" notification*</span></span>

![使用 SurfaceError 例程显示错误](../images/pj15-create-simple-agave-surface-error.png)


## <a name="see-also"></a><span data-ttu-id="e5014-296">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e5014-296">See also</span></span>

- [<span data-ttu-id="e5014-297">Project 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="e5014-297">Task pane add-ins for Project</span></span>](../project/project-add-ins.md)
- [<span data-ttu-id="e5014-298">了解加载项的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e5014-298">Understanding the JavaScript API for add-ins</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="e5014-299">适用于 Office 加载项的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e5014-299">JavaScript API for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="e5014-300">Office 加载项清单的架构参考 (v1.1)</span><span class="sxs-lookup"><span data-stu-id="e5014-300">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)     
- [<span data-ttu-id="e5014-301">Project 2013 SDK 下载</span><span class="sxs-lookup"><span data-stu-id="e5014-301">Project 2013 SDK download</span></span>](https://www.microsoft.com/download/details.aspx?id=30435%20)
    
