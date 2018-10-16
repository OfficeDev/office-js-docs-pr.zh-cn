---
title: OneNote JavaScript API 编程概述
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 557fd1807d860960e7d34587d8ad685c15a883fb
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506271"
---
# <a name="onenote-javascript-api-programming-overview"></a>OneNote JavaScript API 编程概述

OneNote 引入了适用于 OneNote Online 加载项的 JavaScript API。你可以创建任务窗格加载项、内容加载项和与 OneNote 对象交互并连接到 Web 服务或其他基于 Web 的资源的加载项命令。

> [!NOTE]
> 如果计划将加载项[发布](../publish/publish.md)到 AppSource 并使其在 Office 体验内可用，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，要通过验证，加载项就必须在所有平台都管用，这些平台支持你所定义的多种方法（欲知详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性页面](../overview/office-add-in-availability.md)）。

## <a name="components-of-an-office-add-in"></a>Office 加载项组件

加载项由两个基本组件组成：

- 由网页和任何所需的 JavaScript、CSS 或其他文件组成的**Web 应用程序**。这些文件托管在 Web 服务器或 Web 托管服务上，例如 Microsoft Azure。在 OneNote Online 中，Web 应用程序在浏览器控件或 iframe 中显示。
    
- 指定加载项网页的 URL 和适用于加载项的任何访问要求、设置和功能的 **XML 清单**。此文件存储在客户端上。OneNote 加载项使用与其他 Office 加载项相同的[清单](../develop/add-in-manifests.md)格式。

**Office 加载项 = 清单 + 网页**

![Office 加载项包含清单和网页](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>使用 JavaScript API

加载项使用主机应用程序的运行时上下文以访问 JavaScript API。API 有两层： 

- 适用于 OneNote 特定操作的**丰富 API**，通过 **Application** 对象访问。
- 跨 Office 应用程序分享的**公用 API**，通过 **Document** 对象访问。

### <a name="accessing-the-rich-api-through-the-application-object"></a>通过 *Application* 对象访问丰富 API

使用 **Application** 对象访问 OneNote 对象，例如 **Notebook**、**Section** 和 **Page**。通过丰富 API，可以在代理对象上运行批处理操作。基本流程类似如下： 

1. 通过上下文获取应用程序实例。

2. 创建表示你想要使用的 OneNote 对象的代理。通过读取和写入代理对象的属性并调用其方法，你可以与其同步交互。 

3. 调用代理上的 **load** 以使用参数中指定的属性值进行填充。此调用将添加到命令队列中。

   > [!NOTE]
   > API 方法调用（如 `context.application.getActiveSection().pages;`）也会添加到队列中。

4. 调用 **context.sync** 以按命令已排队的顺序运行所有队列中的命令。这将同步你正在运行的脚本和真实对象之间的状态，并通过检索用于你脚本中的已加载的 OneNote 对象的属性实现。你可以使用返回的 promise 对象以链接其他操作。

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

可以在 [API 参考](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js) 中找到所支持的 OneNote 对象和操作。

### <a name="accessing-the-common-api-through-the-document-object"></a>通过 *Document* 对象访问公用 API

使用 **Document** 对象访问公用 API，例如 [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) 和 [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) 方法。 


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
OneNote 加载项仅支持以下公用 API：

| API | 注释 |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) | 仅限 **Office.CoercionType.Text** 和 **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) | 仅限 **Office.CoercionType.Text**、**Office.CoercionType.Image** 和 **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(name);](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) | 仅内容加载项支持设置 | 
| [Office.context.document.settings.set(name, value);](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) | 仅内容加载项支持设置 | 
| [Office.EventType.DocumentSelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) ||

一般情况下，只能使用公用 API 执行在丰富 API 中不支持的操作。要了解有关使用公用 API 的详细信息，请参阅 Office 加载项[文档](../overview/office-add-ins.md)和[引用](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)。


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>OneNote 对象模型图 
下图表示了 OneNote JavaScript API 中当前可用的内容。

  ![OneNote 对象模型图](../images/onenote-om.png)


## <a name="see-also"></a>另请参阅

- [生成首个 OneNote 加载项](onenote-add-ins-getting-started.md)
- [OneNote JavaScript API 参考](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference?view=office-js)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 外接程序平台概述](../overview/office-add-ins.md)
