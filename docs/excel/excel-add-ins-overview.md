---
title: Excel 加载项概述
description: 使用 Excel 加载项，可以跨多个平台（包括 Windows、Mac、iPad 和浏览器）扩展 Excel 应用程序功能。
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 18484c9f6745399b0fdf50da466004e80beca44e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938574"
---
# <a name="excel-add-ins-overview"></a>Excel 加载项概述

使用 Excel 加载项，可以跨多个平台（包括 Windows、Mac、iPad 和浏览器）扩展 Excel 应用程序功能。在工作簿内使用 Excel 加载项，可以：

- 与 Excel 对象交互、读取和写入 Excel 数据。
- 使用基于 Web 的任务窗格或内容窗格扩展功能
- 添加自定义功能区按钮或上下文菜单项
- 添加自定义函数
- 使用对话框窗口提供更丰富的交互

Office 加载项平台提供框架和 Office.js JavaScript API，使你能够创建和运行 Excel 加载项。通过使用 Office 加载项平台创建 Excel 加载项，可以获得以下好处。

- **跨平台支持**：Excel 加载项在 Office 网页版、Windows 版 Office、Mac 版 Office 和 iPad 版 Office中运行。
- **集中式部署**：管理员可以在整个组织内为用户快速而轻松地部署 Excel 加载项。
- **使用标准 Web 技术**：使用熟悉的 Web 技术（如 HTML、CSS 和 JavaScript）创建 Excel 加载项。
- **通过 AppSource 分发**：将 Excel 加载项发布到 [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d)，供广大受众使用。

> [!NOTE]
> Excel 加载项不同于 COM 和 VSTO 加载项，后者是旧 Office 集成解决方案，只能在 Windows 版 Office 上运行。与 COM 加载项不同的是，Excel 加载项不需要你在用户设备上，或在 Excel 中安装任何代码。

## <a name="components-of-an-excel-add-in"></a>Excel 加载项的组件

Excel 加载项包括两个基本组件：Web 应用程序和称为“清单文件”的配置文件。

Web 应用程序使用 [Office JavaScript API](../reference/javascript-api-for-office.md) 与 Excel 中的对象进行交互，并且还有助于与在线资源进行交互。例如，加载项可以执行下列任意任务。

- 创建、读取、更新和删除工作簿中的数据（工作表、区域、表、图表、已命名项等）。
- 使用标准 OAuth 2.0 流通过在线服务执行用户身份验证。
- 向 Microsoft Graph 或任何其他 API 发出 API 请求。

Web 应用程序可以托管在任何 Web 服务器上，并且可以使用客户端框架（如 Angular、React、jQuery）或服务器端技术（如 ASP.NET、Node.js、PHP）进行构建。

[清单](../develop/add-in-manifests.md)是 XML 配置文件，它定义加载项如何通过指定以下设置和功能与 Office 客户端集成：

- 加载项 Web 应用程序的 URL。
- 加载项的显示名称、说明、ID、版本和默认区域设置。
- 如何将加载项与 Excel 集成，其中包括加载项创建的任何自定义 UI（功能区按钮、上下文菜单等）。
- 加载项所需的权限，如对文档执行读取和写入操作。

若要让最终用户能够安装和使用 Excel 加载项，必须将它的清单发布到 AppSource 或加载项目录。 要详细了解如何发布到 AppSource，请参阅[将解决方案发布到 AppSource 和 Office 中](/office/dev/store/submit-to-appsource-via-partner-center)。

## <a name="capabilities-of-an-excel-add-in"></a>Excel 加载项的功能

除了能够与工作簿内容进行交互外，Excel 加载项还可以添加自定义功能区按钮或菜单命令、插入任务窗格、添加自定义函数、打开对话框，甚至还能在工作表中嵌入基于 Web 的丰富对象（如图表或交互式可视化效果）。

### <a name="add-in-commands"></a>加载项命令

加载项命令是能够扩展 Excel UI，并在加载项中启动操作的 UI 元素。加载项命令可用于在功能区中添加按钮，也可用于向 Excel 上下文菜单中添加项。选择加载项命令后，用户便启动操作，如运行 JavaScript 代码，或在任务窗格中显示加载项页面。

**加载项命令**

![Excel 中的加载项命令。](../images/excel-add-in-commands-script-lab.png)

有关命令功能、受支持的平台和开发加载项命令第最佳做法的详细信息，请参阅[适用于 Excel、Word 和 Powerpoint 的加载项命令](../design/add-in-commands.md)。

### <a name="task-panes"></a>任务窗格

任务窗格是接口图面，通常出现在 Excel 中窗口的右侧。使用任务窗格，用户可以访问接口控件，以运行代码来修改 Excel 文档，或显示数据源中的数据。

**任务窗格**

![Excel 中的任务窗格加载项。](../images/excel-add-in-task-pane-insights.png)

有关任务窗格的详细信息，请参阅 [Office 加载项中的任务窗格](../design/task-pane-add-ins.md)。有关在 Excel 中实现任务窗格的示例，请参阅 [Excel 加载项 JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)。

### <a name="custom-functions"></a>自定义函数

开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。

**自定义函数**

![显示最终用户插入 MYFUNCTION 的动画图像。将 SPHEREVOLUME 自定义函数放入 Excel 工作表的单元格中。](../images/SphereVolumeNew.gif)

有关自定义函数的详细信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。

### <a name="dialog-boxes"></a>对话框

对话框是浮动在活动的 Excel 应用程序窗口之上的界面。 可以将对话框用于以下任务，如显示无法直接在任务窗格中打开的登录页、请求用户确认操作，或托管如果局限在任务窗格中可能过小的视频。 若要在 Excel 加载项中打开对话框，请使用[对话框 API](/javascript/api/office/office.ui)。

**对话框**

![Excel 中的加载项对话框。](../images/excel-add-in-dialog-choose-number.png)

有关对话框和对话框 API 的详细信息，请参阅 [Office 加载项中的对话框](../design/dialog-boxes.md)和[在 Office 加载项中使用对话框 API](../develop/dialog-api-in-office-add-ins.md)。

### <a name="content-add-ins"></a>内容加载项

内容加载项是可以直接嵌入到 Excel 文档中的图面。 可以使用内容加载项在工作表中嵌入基于 Web 的丰富对象，如图表、数据可视化效果或媒体，或为用户提供对界面控件的访问权限，这些控件运行代码以修改 Excel 文档，或显示来自数据源的数据。 在你要将功能直接嵌入文档时，请使用内容加载项。

**内容加载项**

![Excel 中的内容加载项。](../images/excel-add-in-content-map.png)

有关内容加载项的详细信息，请参阅 [Office 内容加载项](../design/content-add-ins.md)。有关在 Excel 中实现内容加载项的示例，请参阅 GitHub 中的 [ Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。

## <a name="javascript-apis-to-interact-with-workbook-content"></a>要与工作簿内容交互的 JavaScript API

Excel 加载项通过使用 [Office JavaScript API](../reference/javascript-api-for-office.md) 与 Excel 中的对象进行交互，JavaScript API 包括两个 JavaScript 对象模型：

- **Excel JavaScript API**：[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 随 Office 2016 引入，提供强类型的 Excel 对象，可用于访问工作表、区域、表、图表等。

- **通用 API**：通用 API 随 Office 2013 引入，使用它可以访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。 由于通用 API 确实为 Excel 交互提供了有限的功能，如果加载项需要在 Excel 2013 上运行，则可以使用它。

## <a name="next-steps"></a>后续步骤

通过[创建第一个 Excel 加载项](../quickstarts/excel-quickstart-jquery.md)开始使用。 接下来，请详细了解与生成 Excel 加载项有关的[核心概念](excel-add-ins-core-concepts.md)。

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
- [开发 Office 加载项](../develop/develop-overview.md)
- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
