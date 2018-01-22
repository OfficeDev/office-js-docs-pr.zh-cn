# <a name="onenote-javascript-api-programming-overview"></a>OneNote JavaScript API 编程概述

OneNote 引入了适用于 OneNote Online 外接程序的 JavaScript API。你可以创建任务窗格外接程序、内容外接程序和与 OneNote 对象交互并连接至 Web 服务或其他基于 Web 的资源的外接程序命令。

>
  **注意：**生成外接程序时，如果计划将外接程序[发布](../publish/publish.md)到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。

## <a name="components-of-an-office-add-in"></a>Office 加载项的组件

加载项由两个基本部分组成：

- 包含网页和所有相应 JavaScript、CSS 或其他文件的 **Web 应用程序**。 这些文件托管在 Web 服务器或 Web 托管服务上，例如 Microsoft Azure。 在 OneNote Online 中，Web 应用程序在浏览器控件或 iframe 中显示。
    
- 指定加载项网页 URL 和适用于加载项的任何访问要求、设置和功能的 **XML 清单**。 此文件存储在客户端上。 OneNote 加载项使用的[清单](https://dev.office.com/docs/add-ins/overview/add-in-manifests)格式与其他 Office 加载项相同。

**Office 加载项 = 清单 + 网页**

![Office 外接程序包含清单和网页](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>使用 JavaScript API

外接程序使用托管应用程序的运行时上下文以访问 JavaScript API。API 有两层： 

- 用于执行 OneNote 专属操作的**丰富 API**，可通过 **Application** 对象访问。
- 跨 Office 应用程序分享的**通用 API**，通过**Document** 对象访问。

### <a name="accessing-the-rich-api-through-the-application-object"></a>通过 *Application* 对象访问丰富 API。

**Application** 对象可用于访问 OneNote 对象，如 **Notebook**、**Section** 和 **Page**。 通过丰富 API，您可在代理对象上运行批处理操作。 基本流程类似如下： 

1. 从上下文中获取应用程序实例。

2. 创建您想要使用的表示 OneNote 对象的代理。通过读取和写入代理对象的属性和调用其方法，您可以与其同步交互。 

3. 对代理调用 **load**，使用参数中指定的属性值填充它。 此调用会添加到命令队列中。

    > **注意**：API 方法调用（如 `context.application.getActiveSection().pages;`）也会添加到队列中。

4. 调用 **context.sync**，按排队顺序运行所有已排队命令。 这将同步您正在运行的脚本和真实对象之间的状态，并通过检索已加载的用于您的脚本的 OneNote 对象的属性实现。 您可以使用返回的 promise 对象以链接其他操作。

例如： 

```
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

您可以在 [API 参考](../../reference/onenote/onenote-add-ins-javascript-reference.md) 中找到受支持的 OneNote 对象和操作。

### <a name="accessing-the-common-api-through-the-document-object"></a>通过 *Document* 对象访问通用 API

使用 **Document** 对象以访问通用 API，例如 [getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync) 和 [setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) 方法。 

例如：  

```
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
OneNote 外接程序仅支持以下通用 API：

| API | 注释 |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142294.aspx) | 仅限 **Office.CoercionType.Text** 和 **Office.CoercionType.Matrix** |
| [Office.context.document.setSelectedDataAsync](https://msdn.microsoft.com/en-us/library/office/fp142145.aspx) | 仅限 **Office.CoercionType.Text**、**Office.CoercionType.Image** 和 **Office.CoercionType.Html** | 
| [var mySetting = Office.context.document.settings.get(name);](https://msdn.microsoft.com/en-us/library/office/fp142180.aspx) | 设置仅受内容外接程序支持 | 
| [Office.context.document.settings.set(name, value);](https://msdn.microsoft.com/en-us/library/office/fp161063.aspx) | 设置仅受内容外接程序支持 | 
| [Office.EventType.DocumentSelectionChanged](https://dev.office.com/reference/add-ins/shared/document.selectionchanged.event) ||

一般情况下，仅使用通用 API 执行丰富 API 不支持的操作。 若要详细了解如何使用通用 API，请参阅 Office 加载项[文档](https://dev.office.com/docs/add-ins/overview/office-add-ins)和[参考](https://dev.office.com/reference/add-ins/javascript-api-for-office)。


<a name="om-diagram"></a>
## <a name="onenote-object-model-diagram"></a>OneNote 对象模型图 
下图表示了 OneNote JavaScript API 中当前可用的内容。

  ![OneNote 对象模型图](../images/onenote-om.png)


## <a name="additional-resources"></a>其他资源

- [生成第一个 OneNote 外接程序](onenote-add-ins-getting-started.md)
- [OneNote JavaScript API 参考](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 外接程序平台概述](https://dev.office.com/docs/add-ins/overview/office-add-ins)
