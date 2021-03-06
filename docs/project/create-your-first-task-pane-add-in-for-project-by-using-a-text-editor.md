---
title: 使用文本编辑器为 Microsoft Project 创建首个任务窗格加载项
description: 使用适用于 Office 加载项的 Yeoman 生成器，为 Project Standard 2013、Project Professional 2013 或更高版本创建任务窗格外接程序。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: 4397ad2bd8e44d0962512c2a0905aa1f99a1d411
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159435"
---
# <a name="create-your-first-task-pane-add-in-for-microsoft-project-by-using-a-text-editor"></a>使用文本编辑器为 Microsoft Project 创建首个任务窗格加载项

您可以使用 Office 外接程序的 Yeoman 生成器，为 Project Standard 2013、Project Professional 2013 或更高版本创建任务窗格外接程序。本文介绍如何创建一个简单的外接程序，该外接程序使用指向文件共享上的 HTML 文件的 XML 清单。 项目 OM 测试示例外接程序测试一些使用外接程序的对象模型的 JavaScript 函数。使用 Project 中的**信任中心**注册包含清单文件的文件共享后，可以从功能区上的 "**项目**" 选项卡中打开任务窗格外接程序。 （本文中的示例代码基于 Microsoft Corporation 的 Arvind lyer 所做的测试应用程序。）

Project 使用与其他 Office 客户端相同的加载项清单架构，以及许多相同的 JavaScript API。 Project 2013 SDK 下载的 `Samples\Apps` 子目录中提供了本文所述的加载项的完整代码。

“Project OM 测试”示例加载项可以获取任务的 GUID，以及应用和有效项目的属性。 如果 Project Professional 2013 打开 SharePoint 库中的项目，加载项可以显示项目的 URL。 

[Project 2013 SDK 下载内容](https://www.microsoft.com/download/details.aspx?id=30435%20)包含完整源代码。提取和安装 Project2013SDK.msi 文件中的 SDK 和示例时，请在 `\Samples\Apps\Copy_to_AppManifests_FileShare` 子目录中查找清单文件，并在 `\Samples\Apps\Copy_to_AppSource_FileShare` 子目录中查找源代码。 

JSOMCall.html 示例使用 office.js 文件和 project-15.js 文件中包含的 JavaScript 函数。 可以使用相应的调试文件（office.debug.js 和 project-15.debug.js）检查这些函数。

有关在 Office 外接程序中使用 JavaScript 的简介，请参阅[了解 Office JAVASCRIPT API](../develop/understanding-the-javascript-api-for-office.md)。

## <a name="procedure-1-to-create-the-add-in-manifest-file"></a>过程 1. 创建加载项清单文件

在本地目录中创建一个 XML 文件。 XML 文件包含 `OfficeApp` 元素和子元素，这些元素在[Office 外接程序 XML 清单](../develop/add-in-manifests.md)中进行了介绍。 例如，创建一个名为 JSOM_SimpleOMCalls.xml 的文件，其中包含以下 XML （更改该元素的 GUID 值 `Id` ）。

```XML
<?xml version="1.0" encoding="utf-8"?>
   <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
              xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
              xsi:type="TaskPaneApp">
     <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
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
     <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

对于 Project， `OfficeApp` 元素必须包含 `xsi:type="TaskPaneApp"` 属性值。 `Id`元素是一个 GUID。 `SourceLocation`该值必须为外接程序 HTML 源文件或在任务窗格中运行的 web 应用程序的文件共享路径或 SHAREPOINT URL。 有关清单文件中其他元素的解释，请参阅 [Task pane add-ins for Project](../project/project-add-ins.md)。

过程 2 演示如何创建 JSOM_SimpleOMCalls.XML 清单为 Project 测试加载项指定的 HTML 文件。HTML 文件中指定的按钮调用相关 JavaScript 函数。可以在 HTML 文件内添加 JavaScript 函数，或将它们放在一个单独的 .js 文件中。

## <a name="procedure-2-to-create-the-source-files-for-the-project-om-test-add-in"></a>过程 2. 创建 Project OM Test 加载项的源文件

1. 创建一个 HTML 文件，该文件的名称由 `SourceLocation` JSOM_SimpleOMCalls.xml 清单中的元素指定。 

   例如，在 `C:\Project\AppSource` 目录中创建 theJSOMCall.html 文件。 虽然您可以使用简单的文本编辑器来创建源文件，但使用 Visual Studio Code 这样的工具（例如 HTML 和 JavaScript）会更简单，并且具有其他编辑辅助功能。 如果还未执行 [Project 任务窗格加载项](../project/project-add-ins.md)所述的必应搜索示例，过程 3 将演示如何创建清单指定的 `\\ServerName\AppSource` 文件共享。

   JSOMCall.html 文件为 AJAX 功能使用通用 MicrosoftAjax.js 文件，在 Office 2013 应用程序中使用外接程序功能的 Office.js 文件。

    ```HTML
    <!DOCTYPE html>
    <html>
        <head>
            <title>Project OM Sample Code</title>
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <script type="text/javascript" src="MicrosoftAjax.js"></script>

            <!-- Use the CDN reference to office.js when deploying your add-in. -->
            <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
            <script type="text/javascript" src="office.js"></script>
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

   `textarea`元素指定一个显示 JavaScript 函数结果的文本框。

   > [!NOTE]
   > 为了让“Project OM 测试”示例能够正常运行，请将 Project 2013 SDK 下载内容中的下列文件复制到 JSOMCall.html 文件所在的相同目录：Office.js、Project-15.js 和 MicrosoftAjax.js。

   第 2 步为“Project OM 测试”示例加载项使用的特定函数添加 JSOM_Sample.js 文件。在后续步骤中，将为调用 JavaScript 函数的按钮添加其他 HTML 元素。

2. 在 JSOMCall.html 文件所在的相同目录中，创建 JavaScript 文件 JSOM_Sample.js。 

   下面的代码使用 Office.js 文件中的函数来获取应用程序上下文和文档信息。 该 `text` 对象是 `textarea` HTML 文件中控件的 ID。

   ** \_ ProjDoc**变量是使用对象初始化的 `ProjectDocument` 。 代码包括一些简单的错误处理函数，以及 `getContextValues` 获取应用程序上下文和项目文档上下文属性的函数。 有关 Project 的 JavaScript 对象模型的详细信息，请参阅 [适用于 Office 的 JavaScript API](../reference/javascript-api-for-office.md)。


    ```js
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

   有关 Office.debug.js 文件中的函数的信息，请参阅[Office JAVASCRIPT API](../reference/javascript-api-for-office.md)。 例如，函数将 `getDocumentUrl` 获取打开项目的 URL 或文件路径。

3. 添加调用 Office.js 和 Project-15.js 中异步函数的 JavaScript 函数来获取选定数据：

   - 例如， `getSelectedDataAsync` 是 Office.js 中的一个常规函数，用于获取选定数据的无格式文本。 有关详细信息，请参阅 [AsyncResult 对象](/javascript/api/office/office.asyncresult)。

   - `getSelectedTaskAsync`Project-15.js 中的函数获取所选任务的 GUID。 同样，该 `getSelectedResourceAsync` 函数也会获取所选资源的 GUID。 如果在未选定任务或资源时调用这些函数，函数将显示未定义错误。

   - `getTaskAsync`函数将获取任务名称和已分配资源的名称。 如果任务位于同步的 SharePoint 任务列表中，则 `getTaskAsync` 获取 sharepoint 列表中的任务 id; 否则，sharepoint 任务 id 为0。

     > [!NOTE]
     > 出于演示目的，此示例代码包括一个错误。 如果未 `taskGuid` 定义，则 `getTaskAsync` 函数错误。 如果你获取了一个有效的任务 GUID，然后选择了其他任务，该函数将获取函数对其进行操作的 `getTaskAsync` 最近任务的数据 `getSelectedTaskAsync` 。
  
   - `getTaskFields`、 `getResourceFields` 和 `getProjectFields` 是调用、或多次的本地函数， `getTaskFieldAsync` `getResourceFieldAsync` `getProjectFieldAsync` 用于获取任务或资源的指定字段。 在 project-15.debug.js 文件中， `ProjectTaskFields` 枚举和 `ProjectResourceFields` 枚举显示受支持的字段。

   - `getSelectedViewAsync`函数获取视图的类型（在 project-15.debug.js 中的 `ProjectViewTypes` 枚举中定义）和视图的名称。

   - 如果项目与 SharePoint 任务列表同步，则该函数将 `getWSSUrlAsync` 获取 "任务" 列表的 URL 和名称。 如果项目未与 SharePoint 任务列表同步，则 `getWSSUrlAsync` 函数错误。

     > [!NOTE]
     > 若要获取任务列表的 SharePoint URL 和名称，建议您将 `getProjectFieldAsync` 函数与 `WSSUrl` `WSSList` [ProjectProjectFields](/javascript/api/office/office.projectprojectfields)枚举中的和常量结合使用。

   以下代码的每个函数中都包含由 `function (asyncResult)` 指定的匿名函数，该函数是获取异步结果的回叫。你可以使用命名函数，而不是匿名函数，前者有助于实现复杂外接程序的可维护性。

    ```js
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

4. 添加 JavaScript 事件处理程序回调和函数，以注册任务选择、资源选择和视图选择更改事件处理程序，并注销事件处理程序。 `manageEventHandlerAsync`函数根据_操作_参数添加或删除指定的事件处理程序。 操作可以是 `addHandlerAsync` 或 `removeHandlerAsync` 。

   `manageTaskEventHandler`、 `manageResourceEventHandler` 和 `manageViewEventHandler` 函数可以添加或删除由_docMethod_参数指定的事件处理程序。

    ```js
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

5. 对于 HTML 文档正文，添加调用 JavaScript 函数的按钮进行测试。 例如，在 `div` 常见 JSOM API 的元素中，添加一个调用 general 函数的输入按钮 `getSelectedDataAsync` 。

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

6. 为 `div` 特定于项目的任务函数和事件添加包含按钮的节 `TaskSelectionChanged` 。

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

7. 添加 `div` 包含资源方法和事件、查看方法和事件、项目属性和上下文属性的按钮的节

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

8. 若要设置按钮元素的格式，请添加 CSS `style` 元素。 例如，将以下项添加为元素的子级 `head` 。

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

过程 3 显示如何安装和使用 Project OM Test 加载项功能。

## <a name="procedure-3-to-install-and-use-the-project-om-test-add-in"></a>过程 3. 安装和使用 Project OM Test 加载项

1. 为包含 JSOM_SimpleOMCalls.XML 清单的目录创建一个文件共享。可以在本地计算机或可通过网络访问的远程计算机上创建该文件共享。例如，如果清单位于本地计算机上的 `C:\Project\AppManifests` 目录中，则运行以下命令：

    `Net share AppManifests=C:\Project\AppManifests`

2. 为包含 Project OM Test 加载项的 HTML 和 JavaScript 文件的目录创建一个文件共享。确保文件共享路径与在 JSOM_SimpleOMCalls.xml 清单中指定的路径匹配。例如，如果文件位于本地计算机上的 `C:\Project\AppSource` 目录中，则运行以下命令：

    `net share AppSource=C:\Project\AppSource`

3. 在 Project 中，打开“Project 选项”对话框****，选择“信任中心”****，然后选择“信任中心设置”****。

   [Project 任务窗格加载项](../project/project-add-ins.md)中还介绍了加载项注册过程和其他详细信息。

4. 在“信任中心”**** 对话框的左侧窗格中，选择“受信任的加载项目录”****。

5. 如果您已经 `\\ServerName\AppManifests` 为 Bing 搜索加载项添加了路径，请跳过此步骤。 否则，在 "**受信任的加载项目录**" 窗格中，在 `\\ServerName\AppManifests` "**目录 Url** " 文本框中添加路径，选择 "**添加目录**"，启用网络共享作为默认源（请参阅图1），然后选择 **"确定"**。

   *图 1.为外接程序清单添加网络文件共享*

   ![为应用程序清单添加网络文件共享](../images/pj15-create-simple-agave-manage-catalogs.png)

6. 添加新的加载项或更改源代码后，重新启动 Project。在“Project”**** 功能区上，选择“Office 加载项”**** 下拉菜单，然后选择“查看全部”****。在“插入加载项”**** 对话框中，选择“共享文件夹”****（见图 2），选择“Project OM Test”****，然后选择“插入”****。Project OM Test 加载项将在任务窗格中启动。

   *图 2.启动文件共享上的 Project OM Test 外接程序*

   ![插入应用](../images/pj15-create-simple-agave-start-agave-app.png)

7. 在 Project 中，创建并保存具有至少两个任务的简单项目。 例如，创建名为 T1、T2 的任务和名为 M1 的里程碑，然后将任务工期和前置任务设置为与图 3 中的类似。 选择功能区上的“PROJECT”**** 选项卡，选择任务 T2 的整个行，然后在任务窗格中选择“getSelectedDataAsync”**** 按钮。 图 3 显示在 **Project OM Test** 外接程序的文本框中选择的数据。

   *图 3.使用 Project OM Test 外接程序*

   ![使用 Project OM Test 应用](../images/pj15-create-simple-agave-project-om-test.png)

8. 选择第一项任务的“工期”**** 列中的单元格，然后选择“Project OM Test”加载项中的“getSelectedDataAsync”按钮********。 `getSelectedDataAsync`函数将设置要显示的文本框值 `2 days` 。 

9. 选择所有三项任务的三个“工期”**** 单元格。 `getSelectedDataAsync`函数返回在不同行中选定的单元格的以分号分隔的文本值，例如 `2 days;4 days;0 days` 。

   `getSelectedDataAsync`函数返回在行中选定的单元格的逗号分隔的文本值。 有关图 3 中的示例，选中任务 T2 的整行。 选择时 `getSelectedDataAsync` ，文本框将显示以下内容：`,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`

   “指标”**** 列和“资源名称”**** 列均为空，因此文本数组显示这些列为空值。 `<NA>` 值代表“**添加新列**”单元格。

10. 选择任务 T2 行中的任何单元格，或任务 T2 的整行，然后选择“getSelectedTaskAsync”****。 文本框显示任务 GUID 值，例如，`{25D3E03B-9A7D-E111-92FC-00155D3BA208}`。 Project 将该值存储在 `taskGuid` **项目 OM 测试**外接程序的全局变量中。

11. 选择 "" `getTaskAsync` 。 如果 `taskGuid` 变量包含任务 T2 的 GUID，则文本框将显示任务信息。 **ResourceNames** 值为空。

    创建两个本地资源 R1 andR2，并将其分配给每个50% 的任务 T2，然后再次选择 " **getTaskAsync** "。 文本框中的结果包含资源信息。 如果任务位于同步的 SharePoint 任务列表中，那么结果还会包含 SharePoint 任务 ID。

    - 任务名称：`T2`
    - GUID：`{25D3E03B-9A7D-E111-92FC-00155D3BA208}`
    - WSS ID：`0`
    - ResourceNames: `R1[50%],R2[50%]`

12. 选择 "**获取任务域**" 按钮。 `getTaskFields`函数对 `getTaskfieldAsync` 任务名称、索引、开始日期、工期、优先级和任务备注多次调用函数。

    - 名称：`T2`
    - ID：`2`
    - 开始日期：`Thu 6/14/12`
    - 持续时间：`4d`
    - 优先级：`500`
    - 备注：此为任务 T2 的备注。 仅为测试备注。 若为实际备注，应包含一些真实信息。

13. 选择“getWSSUrlAsync”**** 按钮。如果项目是以下类型之一，结果中显示任务列表 URL 和名称。

    - 导入到 Project Server 的 SharePoint 任务列表。
    - 导入到 Project Professional，然后保存回 SharePoint（未使用 Project Server）的 SharePoint 任务列表。

    > [!NOTE]
    > 如果 Project Professional 安装在 Windows Server 计算机上，则为了能够将项目保存回 SharePoint，您可使用“服务器管理器”**** 添加“桌面体验”**** 功能。

    如果项目是本地项目，或者如果您使用 Project Professional 打开由 Project Server 管理的项目，则该 `getWSSUrlAsync` 方法将显示未定义的错误。

    - SharePoint URL：`http://ServerName`
    - 列表名称：`Test task list`

14. 选择 " **TaskSelectionChanged 事件**" 部分中的 "**添加**" 按钮，它调用 `manageTaskEventHandler` 函数以注册任务选择更改事件，并 `In onComplete function for addHandlerAsync Status: succeeded` 在文本框中返回。 选择一个不同的任务；文本框显示 `In task selection changed event handler`，它是任务选择更改事件的回调函数的输出。 选择“删除”**** 按钮取消注册事件处理程序。

15. 要使用资源方法，可先选择一个视图，如“资源工作表”，****“资源使用状况”**** 或“资源窗体”****，然后在该视图中选择一个资源。 选择**getSelectedResourceAsync**以初始化**resourceGuid**变量，然后选择 "**获取资源域**" 以对 `getResourceFieldAsync` 资源属性多次调用。 还可以添加或删除资源选择更改事件处理程序。

    - 资源名称：`R1`
    - 成本：`$800.00`
    - 标准费率：`$50.00/h`
    - 实际成本：`$0.00`
    - 实际工时：`0h`
    - 单位：`100%`

16. 选择 " **getSelectedViewAsync** " 以显示活动视图的类型和名称。 还可以添加或删除视图选择更改事件处理程序。 例如，如果 "**资源窗体**" 是活动视图，则该 `getSelectedViewAsync` 函数会在文本框中显示以下内容：

    - 视图类型：`6`
    - 名称：`Resource Form`

17. 选择 "**获取项目域**"，为 `getProjectFieldAsync` 活动项目的不同属性多次调用函数。 如果项目是从 Project Web App 打开的，则该 `getProjectFieldAsync` 函数可以获取 Project Web app 实例的 URL。

    - 项目 GUID：`9845922E-DAB4-E111-8AF3-00155D3BA208`
    - 开始日期：`Tue 6/12/12`
    - 完成日期：`Tue 6/19/12`
    - 货币位数：`2`
    - 货币符号：`$`
    - 符号位置：`0`
    - Project Web App URL：`http://servername/pwa`
  
18. 通过获取**Office.Context.document**对象和对象的属性，选择 "**获取上下文值**" 按钮获取文档的属性和运行外接程序的应用程序的属性 `Office.context.application` 。 例如，如果 Project1.mpp 文件在本地计算机桌面上，则文档 URL 为 `C:\Users\UserAlias\Desktop\Project1.mpp`。 如果 .mpp 文件在 SharePoint 库中，则值为文档的 URL。 如果使用 Project Professional 2013 从 Project Web App 打开一个名为 Project1 的项目，则文档 URL 为 `<>\Project1`。

    - 文档 URL：`<>\Project1`
    - 文档模式：`readWrite`
    - 应用语言：`en-US`
    - 显示语言：`en-US`

19. 编辑源代码后，可以通过关闭然后重新启动 Project 来刷新加载项。在“项目”**** 功能区中，“Ofiice 加载项”**** 下拉列表中保留最近使用的加载项的列表。

## <a name="example"></a>示例

Project 2013 SDK 下载包含 JSOMCall.html 文件、JSOM_Sample.js 文件和相关 Office.js、Office.debug.js、Project-15.js、Project-15.debug.js 文件的完整代码。以下是 JSOMCall.html 文件中的代码。

```HTML
<!DOCTYPE html>
<html>
    <head>
        <title>Project OM Sample Code</title>
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

        <script type="text/javascript" src="MicrosoftAjax.js"></script>

        <!-- Use the CDN reference to office.js when deploying your add-in. -->
        <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
        <script type="text/javascript" src="office.js"></script>
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

## <a name="robust-programming"></a>可靠编程

**PROJECT OM Test**外接程序是一个示例，演示如何在 Project-15.js 和 Office.js 文件中使用 Project 2013 的一些 JavaScript 函数。 此示例仅供测试用，不包括可靠的错误检查。 例如，如果您不选择资源并运行 `getSelectedResourceAsync` 函数，则 `resourceGuid` 不会初始化该变量，并且调用将 `getResourceFieldAsync` 返回一个错误。 对于生产加载项，应检查特定错误并忽略结果，隐藏未应用的功能，或通知用户选择视图并在使用函数前先进行有效选择。

对于一个简单的示例，以下代码中的错误输出包括第 `actionMessage` 一个变量，该变量指定为避免函数中的错误要采取的操作 `getSelectedResourceAsync` 。

```js
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

Project 2013 SDK 下载中的 **HelloProject_OData** 示例包含使用 JQuery 库来显示弹出错误消息的 SurfaceErrors.js 文件。 图 4 显示“toast”通知中的错误消息。

SurfaceErrors.js 文件中的以下代码包括 `throwError` 创建对象的 th 函数 `Toast` 。

```js
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

若要使用 `throwError` 函数，请在 JSOMCall.html 文件中包括 JQuery 库和 SurfaceErrors.js 脚本，然后 `throwError` 在其他 JavaScript 函数（如）中添加调用 `logMethodError` 。

> [!NOTE]
> 部署加载项之前，请将 office.js 引用和 jQuery 引用更改为内容发布网络 (CDN) 引用。CDN 引用可提供最新的版本和更好的性能。

```HTML
<!DOCTYPE html>
<html>
<head>
    <title>Project OM Sample Code</title>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <script type="text/javascript" src="MicrosoftAjax.js"></script>

    <!-- Use the CDN reference to Office.js and jQuery when deploying your add-in. -->
    <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
    <script type="text/javascript" src="office.js"></script>
    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jQuery/jquery-1.9.0.min.js"></script>

    <script type="text/javascript" src="JSOM_Sample.js"></script>
    <script type="text/javascript" src="SurfaceErrors.js"></script>

    <!-- . . . INVALID USE OF SYMBOLS . . . -->
</head>

```

<br/>

```js
function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);

    throwError(methodName + " error", actionMessage);
}
```

<br/>

*图 4：SurfaceErrors.js 文件中的函数可以显示“toast”通知*

![使用 SurfaceError 例程显示错误](../images/pj15-create-simple-agave-surface-error.png)


## <a name="see-also"></a>另请参阅

- [Project 任务窗格加载项](../project/project-add-ins.md)
- [了解外接程序的 JavaScript API](../develop/understanding-the-javascript-api-for-office.md)
- [Office JavaScript API 外接程序](../reference/javascript-api-for-office.md)
- [Office 外接程序清单的架构参考 (v1.1)](../develop/add-in-manifests.md)
- [Project 2013 SDK 下载](https://www.microsoft.com/download/details.aspx?id=30435%20)
