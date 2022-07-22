---
title: Visio JavaScript API 概述
description: Visio JavaScript API 概述。
ms.date: 07/18/2022
ms.prod: visio
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 0743057c2f562485c3edb5d3bd82266c13b7e13f
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958886"
---
# <a name="visio-javascript-api-overview"></a>Visio JavaScript API 概述

可以使用 Visio JavaScript API 在 SharePoint Online 的 *经典* SharePoint 页面中嵌入 Visio 图表。 （此拓展性功能在本地 SharePoint 或 SharePoint 框架页面上不支持。）

嵌入的 Visio 图表是存储在 SharePoint 文档库并在 SharePoint 页面上显示的图表。 若要嵌入 Visio 图表，请在 HTML `<iframe>` 元素中显示它。 然后，可以使用 Visio JavaScript API 以程序化方式处理嵌入的图表。

![SharePoint 页面上 iframe 中的 Visio 图表，以及脚本编辑器 Web 部件。](../images/visio-api-block-diagram.png)

可以使用 Visio JavaScript API 执行以下操作：

- 与页面、形状等 Visio 图表元素进行交互。
- 在 Visio 图表画布上创建视觉标记。
- 为绘图中的鼠标事件编写自定义处理程序。
- 向解决方案公开图表数据，如形状文本、形状数据和超链接。

本文介绍了如何通过结合使用 Visio JavaScript API 和 Visio 网页版来生成 SharePoint Online 解决方案。具体介绍了有关使用 API（如 `EmbeddedSession`、`RequestContext`、JavaScript 代理对象、`sync()`、`Visio.run()` 和 `load()` 方法）的基本概念。下面这些代码示例展示了如何应用这些概念。

## <a name="embeddedsession"></a>EmbeddedSession

EmbeddedSession 对象在浏览器中初始化开发人员框架和 Visio 框架之间的通信。

```js
const session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
session.init().then(function () {
    window.console.log("Session successfully initialized");
});
```

## <a name="visiorunsession-functioncontext--batch-"></a>Visio.run(session, function(context) { batch })

`Visio.run()` 运行一个对 Visio 对象模型执行操作的批处理脚本。 批处理命令包括定义本地 JavaScript 代理对象、在本地和 Visio 对象之间同步状态的 `sync()` 方法以及承诺实现。 `Visio.run()` 中的批处理请求的优势在于，当实现承诺时，在执行期间分配的任何被跟踪的页面对象将会自动释放。

`run` 函数获取会话和 RequestContext 对象，并返回一个承诺（通常就是 `context.sync()` 的结果）。 可以在 `Visio.run()` 之外运行批处理操作。 不过，在这种情况下，需要手动跟踪和管理任何页面对象引用。

## <a name="requestcontext"></a>RequestContext

RequestContext 对象可方便对 Visio 应用程序提出请求。由于开发人员框架和 Visio Web 客户端在两个不同的 iframe 中运行，因此 RequestContext 对象（下一个示例中的上下文）必须能够从开发人员框架访问 Visio 和相关对象（如页面和形状）。

```js
function hideToolbars() {
    Visio.run(session, function(context){
        const app = context.document.application;
        app.showToolbars = false;
        return context.sync().then(function () {
            window.console.log("Toolbars Hidden");
        });
    }).catch(function(error)
    {
        window.console.log("Error: " + error);
    });
};
```

## <a name="proxy-objects"></a>代理对象

在嵌入式会话中声明和使用的 Visio JavaScript 对象是 Visio 文档中真实对象的代理对象。对代理对象执行的所有操作都不会在 Visio 中实现，并且在同步文档状态前，Visio 文档的状态不会在代理对象中实现。运行 `context.sync()` 时将同步文档状态。

例如，本地 JavaScript 对象 getActivePage 声明为引用选定页面。 这可用于将属性和调用方法的设置操作排入队列。 对此类对象执行的操作不会实现，除非运行 `sync()` 方法。

```js
const activePage = context.document.getActivePage();
```

## <a name="sync"></a>sync()

`sync()` 方法通过执行在上下文中排队的指令以及检索用于你代码中的已加载 Office 对象的属性，在 JavaScript 代理对象和 Visio 中的真实对象之间同步状态。此方法返回一个将在同步完成时实现的承诺。

## <a name="load"></a>load()

`load()` 方法用于填充在 JavaScript 层中创建的代理对象。当尝试检索一个对象（如文档）时，将首先在 JavaScript 层创建一个本地代理对象。此类对象可用于对其属性和调用方法的设置进行排队。但为了读取对象属性或关系，需要首先调用 `load()` 和 `sync()` 方法。当调用 `sync()` 方法时，load() 方法将接受需要加载的属性和关系。

下面的示例展示了 `load()` 方法的语法。

```js
object.load(string: properties); //or object.load(array: properties); //or object.load({loadOption});
```

1. **properties** 列出了要加载的属性名称，指定为逗号分隔的字符串或名称数组。 有关详细信息，请参阅每个对象下的 `.load()` 方法。

2. **loadOption** 指定的对象描述了选择、展开、置顶和跳过选项。有关详细信息，请参阅对象加载 [选项](/javascript/api/office/officeextension.loadoption)。

## <a name="example-printing-all-shapes-text-in-active-page"></a>示例：打印活动页中的所有形状文本

下面的示例展示了如何打印数组形状对象的形状文本值。
`Visio.run()` 函数包含一批指令。 在此次批处理期间，将会创建一个代理对象，引用活动文档中的形状。

所有这些命令将在调用 `context.sync()` 时排入队列和运行。 `sync()` 方法返回一个承诺，可用于将其与其他操作关联起来。

```js
Visio.run(session, function (context) {
    const page = context.document.getActivePage();
    const shapes = page.shapes;
    shapes.load();
    return context.sync().then(function () {
        for(let i=0; i<shapes.items.length;i++) {
            let shape = shapes.items[i];
            window.console.log("Shape Text: " + shape.text );
        }
    });
}).catch(function(error) {
    window.console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        window.console.log ("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="error-messages"></a>错误消息

使用包含代码和消息的错误对象返回错误。下表列出了可能发生的错误情况。

| error.code            | error.message |
|-----------------------|----------------------------------------------------------------|
| InvalidArgument       | 自变量无效、缺少或格式不正确。 |
| GeneralException      | 处理请求时出现内部错误。 |
| NotImplemented        | 所请求的功能未实现。  |
| UnsupportedOperation  | 不支持正在尝试的操作。 |
| AccessDenied          | 无法执行所请求的操作。 |
| ItemNotFound          | 所请求的资源不存在。 |

## <a name="get-started"></a>开始使用

可以从本部分中的示例入手。 此示例展示了如何在 Visio 图表中以编程方式显示选定形状的形状文本。 首先，在 SharePoint Online 中创建一个经典页面，或编辑现有页面。 在页面上添加脚本编辑器 Web 部件，复制并粘贴下面的代码。

```HTML
<script src='https://appsforoffice.microsoft.com/embedded/1.0/visio-web-embedded.js' type='text/javascript'></script>

Enter Visio File Url:<br/>
<script language="javascript">
document.write("<input type='text' id='fileUrl' size='120'/>");
document.write("<input type='button' value='InitEmbeddedFrame' onclick='initEmbeddedFrame()' />");
document.write("<br />");
document.write("<input type='button' value='SelectedShapeText' onclick='getSelectedShapeText()' />");
document.write("<textarea id='ResultOutput' style='width:350px;height:60px'> </textarea>");
document.write("<div id='iframeHost' />");

let session; // Global variable to store the session and pass it afterwards in Visio.run()
let textArea;
// Loads the Visio application and Initializes communication between developer frame and Visio online frame
function initEmbeddedFrame() {
    textArea = document.getElementById('ResultOutput');
    let url = document.getElementById('fileUrl').value;
    if (!url) {
        window.alert("File URL should not be empty");
    }
    // APIs are enabled for EmbedView action only.
    url = url.replace("action=view","action=embedview");
    url = url.replace("action=interactivepreview","action=embedview");
    url = url.replace("action=default","action=embedview");
    url = url.replace("action=edit","action=embedview");
  
    session = new OfficeExtension.EmbeddedSession(url, { id: "embed-iframe",container: document.getElementById("iframeHost") });
    return session.init().then(function () {
        // Initialization is successful
        textArea.value  = "Initialization is successful";
    });
}

// Code for getting selected Shape Text using the shapes collection object
function getSelectedShapeText() {
    Visio.run(session, function (context) {
        const page = context.document.getActivePage();
        const shapes = page.shapes;
        shapes.load();
        return context.sync().then(function () {
            textArea.value = "Please select a Shape in the Diagram";
            for(let i=0; i<shapes.items.length;i++) {
                let shape = shapes.items[i];
                if ( shape.select == true) {
                    textArea.value = shape.text;
                    return;
                }
            }
        });
    }).catch(function(error) {
        textArea.value = "Error: ";
        if (error instanceof OfficeExtension.Error) {
            textArea.value += "Debug info: " + JSON.stringify(error.debugInfo);
        }
    });
}
</script>
```

完成此操作之后，只需使用你想要使用的 Visio 图表的 URL。 只需将 Visio 图表上传到 SharePoint Online 并在 Visio 网页版将其打开。 在这里打开嵌入对话框，然后使用以上示例中的嵌入 URL。

![复制嵌入对话框中的 Visio 文件 URL。](../images/Visio-embed-url.png)

如果在编辑模式下使用 Visio 网页版，请通过依次选择“**文件**” > “**共享**” > “**嵌入**”来打开嵌入对话框。 如果在视图模式下使用 Visio 网页版，请通过选择“...”和“**嵌入**”来打开嵌入对话框。

## <a name="visio-javascript-api-reference"></a>Visio JavaScript API 参考

有关 Visio JavaScript API 的详细信息，请参阅 [Visio JavaScript API 参考文档](/javascript/api/visio)。
