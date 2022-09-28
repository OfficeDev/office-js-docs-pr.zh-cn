---
title: OneNote JavaScript API 编程概述
description: 了解有关适用于 OneNote 网页版加载项的 OneNote JavaScript API。
ms.date: 07/18/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: d44a01cf0f676057ca072cff74e2e80057f645f4
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092908"
---
# <a name="onenote-javascript-api-programming-overview"></a>OneNote JavaScript API 编程概述

OneNote introduces a JavaScript API for OneNote add-ins on the web. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a>Office 加载项的组件

加载项由两个基本部分组成：

- A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote on the web, the web application displays in a browser control or iframe.

- An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.

### <a name="office-add-in--manifest--webpage"></a>Office 加载项 = 清单 + 网页

![Office 加载项包含清单和网页。](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>使用 JavaScript API

Add-ins use the runtime context of the Office application to access the JavaScript API. The API has two layers:

- 用于执行 OneNote 专属操作的 **应用程序特定 API**，可通过 `Application` 对象访问。
- 跨 Office 应用程序分享的 **通用 API**，通过 `Document` 对象访问。

### <a name="accessing-the-application-specific-api-through-the-application-object"></a>通过 *Application* 对象访问应用程序特定 API。

Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With application-specific APIs, you run batch operations on proxy objects. The basic flow goes something like this:

1. 从上下文中获取应用程序实例。

2. Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.

3. Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.

   > [!NOTE]
   > API 方法调用（如 `context.application.getActiveSection().pages;`）也会添加到队列中。

4. Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.

例如：

```js
async function getPagesInSection() {
    await OneNote.run(async (context) => {

        // Get the pages in the current section.
        const pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        await context.sync();
            
        // Read the id and title of each page.
        $.each(pages.items, function(index, page) {
            let pageId = page.id;
            let pageTitle = page.title;
            console.log(pageTitle + ': ' + pageId);
        });
    });
}
```

有关详细信息，请参阅[使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，了解 OneNote JavaScript API 中的 `load`/`sync` 模式以及其他常见做法。

可以在 [API 参考](../reference/overview/onenote-add-ins-javascript-reference.md) 中找到受支持的 OneNote 对象和操作。

#### <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API 要求集

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets).

### <a name="accessing-the-common-api-through-the-document-object"></a>通过 *Document* 对象访问通用 API

使用 `Document` 对象以访问通用 API，例如 [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) 和 [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) 方法。

例如：  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            const error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

OneNote 加载项仅支持以下通用 API。

| API | 注释 |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) | 仅 `Office.CoercionType.Text` 和 `Office.CoercionType.Matrix` |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) | 仅 `Office.CoercionType.Text`、`Office.CoercionType.Image` 和 `Office.CoercionType.Html` |
| [const mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#office-office-settings-get-member(1)) | 设置仅受内容外接程序支持 |
| [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#office-office-settings-set-member(1)) | 设置仅受内容外接程序支持 |
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) |*没有。*|

一般情况下，需要使用通用 API 执行应用程序特定 API 不支持的操作。 要详细了解如何使用通用 API，请参阅[常见 JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。

<a name="om-diagram"></a>

## <a name="onenote-object-model-diagram"></a>OneNote 对象模型图

下图表示了 OneNote JavaScript API 中当前可用的内容。

  ![OneNote 对象模型图。](../images/onenote-om.png)

## <a name="see-also"></a>另请参阅

- [开发 Office 加载项](../develop/develop-overview.md)
- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
- [生成首个 OneNote 加载项](../quickstarts/onenote-quickstart.md)
- [OneNote JavaScript API 参考](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 加载项平台概述](../overview/office-add-ins.md)
