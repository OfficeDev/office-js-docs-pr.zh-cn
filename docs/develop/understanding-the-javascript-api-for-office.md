---
title: 了解适用于 Office 的 JavaScript API
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: ccc5122061e267fec875fcbbb5b2083e1b934f9d
ms.sourcegitcommit: 7ecc1dc24bf7488b53117d7a83ad60e952a6f7aa
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/22/2018
ms.locfileid: "22546786"
---
# <a name="understanding-the-javascript-api-for-office"></a>了解适用于 Office 的 JavaScript API

本文提供了有关适用于 Office 的 JavaScript API 的信息以及使用方法。有关参考信息，请参阅 [适用于 Office 的 JavaScript API](https://dev.office.com/reference/add-ins/javascript-api-for-office)。有关将 Visual Studio 项目文件更新到适用于 Office 的 JavaScript API 的最新当前版本的信息，请参阅 [更新适用于 Office 的 JavaScript API 版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。

> [!NOTE]
> 如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>在加载项中引用适用于 Office 的 JavaScript API 库

[适用于 Office 的 JavaScript](https://dev.office.com/reference/add-ins/javascript-api-for-office) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

这将在加载项首次加载时下载并缓存适用于 Office 的 JavaScript API 文件，以确保对特定版本使用 Office.js 及其关联文件的最新实现。

有关 Office.js CDN 的更多详细信息（包括如何处理版本控制和向后兼容性），请参阅[从内容分发网络 (CDN) 引用适用于 Office 的 JavaScript API 库](referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

## <a name="initializing-your-add-in"></a>初始化加载项

**适用于：** 所有加载项类型

Office.js 提供初始化事件，API 完全加载并准备与用户开始交互时会触发该事件。你可以使用 **initialize** 事件处理程序实现常见的外接程序初始化方案，例如，可以提示用户选择 Excel 中的一些单元格，然后插入使用选定值初始化的图表。还可以使用 initialize 事件处理程序初始化外接程序的其他自定义逻辑，例如建立绑定、提示默认外接程序设置值等。

至少，initialize 事件应类似下面的示例：     

```js
Office.initialize = function () { };
```
如果你使用其他 JavaScript 框架，其中包括它们自己的初始化处理程序或测试，那么它们应放置在 Office.initialize 事件内。例如，会对 [JQuery](https://jquery.com) `$(document).ready()` 函数进行以下引用：

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

Office 外接程序中的所有页面需要向 initialize 事件 (**Office.initialize**) 分配一个事件处理程序。如果未能分配一个事件处理程序，则外接程序可能会在启动时出现错误。而且，如果某个用户尝试通过 Office Online Web 客户端（例如 Excel Online、PowerPoint Online 或 Outlook Web App）使用你的外接程序，则外接程序会无法运行。如果无需任何初始化代码，则向 **Office.initialize** 分配的函数的正文可以如同上述第一个示例中一样为空。

若要详细了解加载项初始化时的事件发生顺序，请参阅[加载 DOM 和运行时环境](loading-the-dom-and-runtime-environment.md)。

#### <a name="initialization-reason"></a>初始化原因
Office.initialize 为任务窗格和内容外接程序提供其他“_reason_”参数。此参数可用于确定如何将外接程序添加到当前文档。你可以使用此参数针对首次插入加载项时和加载项已存在于文档中时实施不同的逻辑。 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```
有关详细信息，请参阅 [Office.initialize 事件](https://dev.office.com/reference/add-ins/shared/office.initialize)和 [InitializationReason 枚举](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration)。 

## <a name="office-javascript-api-object-model"></a>Office JavaScript API 对象模型

初始化后，加载项可以与主机（例如 Excel、Outlook）交互。 [Office JavaScript API 对象模型](office-javascript-api-object-model.md)页面有关于特定使用模式的更多详细信息。 [共享 API](https://dev.office.com/reference/add-ins/javascript-api-for-office) 和特定的主机都有详细的参考文档。

## <a name="api-support-matrix"></a>API 支持矩阵


下表总结了各种类型的加载项（内容、任务窗格和 Outlook）支持的 API 和功能，以及使用[适用于 Office 的 JavaScript API v1.1 支持的 1.1 加载项清单架构和功能](update-your-javascript-api-for-office-and-manifest-schema-version.md)指定加载项支持的 Office 主机应用时，可以托管它们的 Office 应用。


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**主机名**|数据库|工作簿|邮箱|演示文稿|文档|项目|
||**支持的****主机应用程序**|Access Web App|Excel、<br/>Excel Online|Outlook、<br/>Outlook Web App、<br/>适用于设备的 OWA|PowerPoint、<br/>PowerPoint Online|Word|项目|
|**支持的外接程序类型**|内容|是|是||是|||
||任务窗格||是||是|是|是|
||Outlook|||是||||
|**支持的 API 功能**|读/写文本||是||是|是|是<br/>（只读）|
||读/写矩阵||是|||是||
||读/写表||是|||是||
||读/写 HTML|||||是||
||读/写<br/>Office Open XML|||||是||
||读取任务、资源、视图和字段属性||||||是|
||选择已更改事件||是|||是||
||获取整个文档||||是|是||
||绑定和绑定事件|是<br/>（仅限完全和部分表格绑定）|是|||是||
||读/写自定义 XML 部分|||||是||
||暂留加载项状态数据（设置）|是<br/>（每主机加载项）|是<br/>（每文档）|是<br/>（每邮箱）|是<br/>（每文档）|是<br/>（每文档）||
||设置更改事件|是|是||是|是||
||获取活动视图模式<br/>和视图更改事件||||是|||
||转到文档中<br/>的相应位置||是||是|是||
||使用规则和 RegEx<br/>执行上下文式激活|||是||||
||读取项目属性|||是||||
||读取用户配置文件|||是||||
||获取附件|||是||||
||获取用户标识令牌|||是||||
||调用 Exchange Web 服务|||是||||
