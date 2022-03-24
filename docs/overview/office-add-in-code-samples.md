---
title: Office 加载项代码示例
description: Office 加载项代码示例列表，可帮助你学习和生成自己的加载项。
ms.date: 02/17/2022
localization_priority: high
ms.openlocfilehash: e727e1df0bfb02eade1133e575234554f7c2b144
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745962"
---
# <a name="office-add-in-code-samples"></a>Office 加载项代码示例

编写这些代码示例的目的是为了帮助你了解如何在开发 Office 加载项时使用各种功能。

## <a name="getting-started"></a>入门

以下示例演示如何仅使用清单、HTML 网页和徽标生成最简单的 Office 加载项。 这些组件是 Office 加载项的基本部分。 有关其他入门信息，请参阅我们的 [快速入门](../quickstarts/excel-quickstart-jquery.md) 和 [教程](/search/?terms=tutorial&scope=Office%20Add-ins)。

- [Excel "Hello world" 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world)
- [Outlook "Hello world" 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)
- [PowerPoint "Hello world" 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world)
- [Word "Hello world" 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world)

<br>

---

---

## <a name="excel"></a>Excel

| 名称                | 说明         |
|:--------------------|:--------------------|
| [在 Teams 中打开](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-open-in-teams) | 在 Microsoft Teams 中新建包含你定义的数据的 Excel 电子表格。|
| [插入外部 Excel 文件并使用 JSON 数据填充](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-insert-file)  | 将外部 Excel 文件中的现有模板插入当前打开的 Excel 工作簿。 然后，使用来自 JSON Web 服务的数据填充模板。 |
| [在功能区上创建自定义上下文选项卡](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs) | 在 Office UI 中的功能区上创建自定义上下文选项卡。 该示例创建一个表，并且当用户将焦点移动到表内时，将显示自定义选项卡。 当用户移出表外时，自定义选项卡将隐藏。 |
| [使用键盘快捷方式执行 Office 加载项操作](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) | 设置使用键盘快捷方式的基本 Excel 加载项项目。 |
| [使用 Web 辅助进程的自定义函数示例](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/web-worker) | 在自定义函数中使用 Web 辅助进程来防止阻止 Office 加载项的 UI。 |
| [脱机时使用存储技术从 Office 加载项访问数据](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin) | 实施 localStorage，以便在用户遇到连接丢失时为 Office 加载项启用有限的功能。 |
| [自定义函数批处理模式](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching)| 将多个调用批处理为单个调用，以减少对远程服务的网络调用数。|

## <a name="outlook"></a>Outlook

| 名称                | 说明         |
|:--------------------|:--------------------|
| [加密附件、处理会议请求与会者以及对约会日期/时间更改做出回应](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-encrypt-attachments) | 当用户添加时，使用基于事件的激活来加密附件。还可以对会议请求中更改的收件人以及会议请求中开始或结束日期或时间的更改使用事件处理。 |
| [使用 Outlook 基于事件的激活标记外部收件人](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external) | 当用户在撰写邮件时更改收件人时，使用基于事件的激活运行 Outlook 加载项。 加载项还使用 `appendOnSendAsync` API 添加免责声明。 |
| [使用 Outlook 基于事件的激活设置签名](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature) | 用户创建新邮件或约会时，基于事件的激活将运行 Outlook 加载项。 即使没有打开任务窗格，加载项也可以响应事件。 它还使用 `setSignatureAsync` API。 |

## <a name="word"></a>Word

| 名称                | 说明         |
|:--------------------|:--------------------|
| [使用 Word 加载项获取、编辑和设置 Word 文档中的 OOXML 内容](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-get-set-edit-openxml) | 此示例展示了如何获取、编辑和设置 Word 文档中的 OOXML 内容。 示例加载项提供了一个暂存区，用于获取自己的内容的 Office Open XML，并测试自己编辑的 Office Open XML 代码片段。|
| [在 Word 加载项中加载和写入 Open XML](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)  | 本示例加载项显示如何通过将 setSelectedDataAsync 方法与 ooxml coercion 类型结合使用，将多种丰富的内容类型添加到 Word 文档。 还可以通过此加载项直接在页面上显示每个示例内容类型的 Office Open XML 标记。 |

<br>

---

---

## <a name="authentication-authorization-and-single-sign-on-sso"></a>身份验证、授权和单一登录 (SSO)

| 名称                | 说明         |
|:--------------------|:--------------------|
| [单一登录 (SSO) 示例 Outlook 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) | 使用 Office 的 SSO 功能向加载项提供 Microsoft Graph 数据的访问权限。|
| [使用 Microsoft Graph 和 Office 加载项中的 msal.js 获取 OneDrive 数据](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React) | 将 Office 加载项构建为一个没有后端的单页应用程序 (SPA)，该应用程序连接到 Microsoft Graph，并访问存储在 OneDrive for Business 中的工作簿以更新电子表格。  |
| [Office 加载项对 Microsoft Graph 的身份验证](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) | 了解如何构建连接到 Microsoft Graph 的 Microsoft Office 加载项，并访问存储在 OneDrive for Business 中工作簿以更新电子表格。。 |
| [Outlook 加载项对 Microsoft Graph 的身份验证](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)。 | 生成连接到 Microsoft Graph 的 Outlook 加载项，并访问存储在 OneDrive for Busines s中的工作簿以撰写新的电子邮件。 |
| [带有 ASP.NET 的单一登录 (SSO) Office 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) | 在 Office.js 中使用 `getAccessToken` API 为加载项提供 Microsoft Graph 数据的访问权限。此示例基于 ASP.NET。 |
| [带有 Node.js 的单一登录 (SSO) Office 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | 在 Office.js 中使用 `getAccessToken` API 为加载项提供 Microsoft Graph 数据的访问权限。此示例基于 Node.js 构建。|

## <a name="shared-javascript-runtime"></a>共享 JavaScript 运行时

| 名称                | 说明         |
|:--------------------|:--------------------|
| [与共享运行时共享全局数据](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-global-state) | 设置使用共享运行时在单个浏览器运行时中运行功能区按钮、任务窗格和自定义函数代码的基本项目。 |
| [管理功能区和任务窗格 UI，并在打开文档时运行代码](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario) | 创建根据加载项状态启用的上下文功能区按钮。 |

<br>

---

---

## <a name="additional-samples"></a>其他示例

| 名称                | 说明         |
|:--------------------|:--------------------|
| [使用共享库将 Visual Studio Tools for Office 加载项迁移到 Office Web 加载项](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/VSTO-shared-code-migration) | 提供从 VSTO 加载项迁移到 Office 加载项时代码重用的策略。 |
| [将 Azure 函数与 Excel 自定义函数集成](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AzureFunction) | 将 Azure Functions 与自定义函数集成，以移动到云或集成其他服务。 |
| [动态 DPI 代码示例](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/dynamic-dpi) | 用于处理 COM、VSTO 和 Office 加载项中 DPI 更改的示例集合。 |

## <a name="next-steps"></a>后续步骤

加入 Microsoft 365 开发人员计划。获取为 Microsoft 365 平台构建解决方案所需的免费沙盒、工具和其他资源。

- [免费开发人员沙盒](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) 获取免费的可续订 90 天 Microsoft 365 E5 开发人员订阅。
- [示例数据包](https://developer.microsoft.com/microsoft-365/dev-program#Sample) 通过安装用户数据和内容来帮助你构建解决方案，从而自动配置你的沙盒。
- [访问专家](https://developer.microsoft.com/microsoft-365/dev-program#Experts) 参与社区活动，以向 Microsoft 365 专家学习。
- [个性化建议](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) 快速从个性化仪表板查找开发人员资源。
