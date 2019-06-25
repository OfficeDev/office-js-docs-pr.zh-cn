---
title: OneNote JavaScript API 编程概述
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 852c68bc9edf370d0eef687fb4869b23d4f59fe4
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128634"
---
# <a name="onenote-javascript-api-programming-overview"></a>OneNote JavaScript API 编程概述

OneNote 引入了适用于 OneNote 网页版加载项的 JavaScript API。 可以创建任务窗格加载项、内容加载项，以及与 OneNote 对象交互并连接到 Web 服务或其他基于 Web 的资源的加载项命令。

> [!NOTE]
> 如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。

## <a name="components-of-an-office-add-in"></a>Office 加载项的组件

加载项由两个基本部分组成：

- 包含网页和所有相应 JavaScript、CSS 或其他文件的 **Web 应用程序**。 这些文件托管在 Web 服务器或 Web 托管服务上，例如 Microsoft Azure。 在 OneNote 网页版中，Web 应用程序在浏览器控件或 iframe 中显示。

- **XML 清单**指定外接程序网页的 URL 和适用于外接程序的任何访问要求、设置和功能。此文件存储在客户端上。OneNote 外接程序使用与其他 Office 外接程序相同的 [清单](../develop/add-in-manifests.md)格式。

**Office 加载项 = 清单 + 网页**

![Office 加载项包含清单和网页](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>使用 JavaScript API

加载项使用托管应用程序的运行时上下文以访问 JavaScript API。API 有两层： 

- 用于执行 OneNote 专属操作的**主机特定 API**，可通过 **Application** 对象访问。
- 跨 Office 应用程序分享的**通用 API**，通过**Document** 对象访问。

### <a name="accessing-the-host-specific-api-through-the-application-object"></a>通过 *Application* 对象访问主机特定 API。

**Application** 对象可用于访问 OneNote 对象，如 **Notebook**、**Section** 和 **Page**。 通过主机特定 API，可在代理对象上运行批处理操作。 基本流程类似如下： 

1. 从上下文中获取应用程序实例。

2. 创建您想要使用的表示 OneNote 对象的代理。通过读取和写入代理对象的属性和调用其方法，您可以与其同步交互。

3. 调用代理上的 **load** 以使用在参数中指定的属性值填充它。此调用将添加至命令队列中。

   > [!NOTE]
   > API 方法调用（如 `context.application.getActiveSection().pages;`）也会添加到队列中。

4. 调用 **context.sync** 以按它们已排队的顺序运行所有排队的命令。这将同步您正在运行的脚本和真实对象之间的状态，并通过检索已加载的用于您的脚本的 OneNote 对象的属性实现。您可以使用返回的 promise 对象以链接其他操作。

例如：

```js
function getPagesInSection() {
    OneNote.run(function (context) {

        // Get the pages in the current section.
        var pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync()
            .then(function () {

                // Read the id and title of each page.
                $.each(pages.items, function(index, page) {
                    var pageId = page.id;
                    var pageTitle = page.title;
                    console.log(pageTitle + ': ' + pageId);
                });
            })
            .catch(function (error) {
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    });
}
```

可以在 [API 参考](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference) 中找到受支持的 OneNote 对象和操作。

### <a name="accessing-the-common-api-through-the-document-object"></a>通过 *Document* 对象访问通用 API

使用 **Document** 对象以访问通用 API，例如 [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 和 [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 方法。 


例如：  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

OneNote 加载项仅支持以下通用 API：

| API | 注释 |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) | 仅限 **Office.CoercionType.Text** 和 **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) | 仅限 **Office.CoercionType.Text**、**Office.CoercionType.Image** 和 **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#get-name-) | 设置仅受内容外接程序支持 | 
| [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#set-name--value-) | 设置仅受内容外接程序支持 | 
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) ||

一般情况下，仅使用通用 API 执行主机特定 API 不支持的操作。 若要详细了解如何使用通用 API，请参阅 Office 加载项[文档](../overview/office-add-ins.md)和[参考](../reference/javascript-api-for-office.md)。


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>OneNote 对象模型图 
下图表示了 OneNote JavaScript API 中当前可用的内容。

  ![OneNote 对象模型图](../images/onenote-om.png)


## <a name="see-also"></a>另请参阅

- [生成首个 OneNote 加载项](../quickstarts/onenote-quickstart.md)
- [OneNote JavaScript API 参考](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 加载项平台概述](../overview/office-add-ins.md)
