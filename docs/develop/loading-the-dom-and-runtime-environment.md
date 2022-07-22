---
title: 加载 DOM 和运行时环境
description: 加载 DOM 和 Office 加载项运行时环境。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: be93b261c8beacdb7b4e8cd08448abf06b14607e
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958683"
---
# <a name="loading-the-dom-and-runtime-environment"></a>加载 DOM 和运行时环境

外接程序在运行自己的自定义逻辑前必须确保 DOM 和 Office 外接程序运行时环境都已加载。

## <a name="startup-of-a-content-or-task-pane-add-in"></a>启动内容或任务窗格加载项

下图显示了在 Excel、PowerPoint、Project 或 Word 中启动内容或任务窗格加载项所涉及的事件流。

![启动内容或任务窗格加载项时的事件流。](../images/office15-app-sdk-loading-dom-agave-runtime.png)

当内容或任务窗格加载项启动时，将发生以下事件。

1. 用户打开已包含加载项的文档，或在文档中插入加载项。

2. Office 客户端应用程序从 AppSource、SharePoint 上的应用目录或它源自的共享文件夹目录中读取外接程序的 XML 清单。

3. Office 客户端应用程序在浏览器控件中打开外接程序的 HTML 页面。

    后面的两个步骤第 4 步和第 5 步以异步方式并行发生。因此，您的加载项代码必须在继续之前确保 DOM 和加载项运行时环境已加载完。

4. 浏览器控件加载 DOM 和 HTML 正文，并调用事件的事件处理程序 `window.onload` 。

5. Office 客户端应用程序加载运行时环境，该环境从内容分发网络 (CDN) 服务器下载和缓存 Office JavaScript API 库文件，然后调用外接程序的事件处理程序来 [初始化](/javascript/api/office#Office_initialize_reason_) [Office](/javascript/api/office) 对象的事件（如果已分配了处理程序）。 此时，它还检查是否已将任何回调 (或链接 `then()` 方法) 传递 (或链接) 到 `Office.onReady` 处理程序。 有关区别`Office.initialize``Office.onReady`的详细信息，请参阅[“初始化加载项](initialize-add-in.md)”。

6. 当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。

## <a name="startup-of-an-outlook-add-in"></a>启动 Outlook 外接程序

下图显示了启动在台式机、平板电脑或智能手机上运行的 Outlook 外接程序所涉及的事件流。

![启动 Outlook 加载项时的事件流。](../images/outlook15-loading-dom-agave-runtime.png)

Outlook 加载项启动时会发生以下事件。

1. 当 Outlook 启动时，Outlook 读取已为用户的电子邮件帐户安装的 Outlook 外接程序的 XML 清单。

2. 用户选择 Outlook 中的一个项目。

3. 如果所选项目满足某个 Outlook 外接程序的激活条件，则 Outlook 将激活该外接程序，并使其按钮在 UI 中可见。

4. 如果用户单击该按钮以启动 Outlook 外接程序，Outlook 将在浏览器控件中打开 HTML 页面。下面两个步骤（步骤 5 和 6）并行发生。

5. 浏览器控件加载 DOM 和 HTML 正文，并调用事件的事件处理程序 `onload` 。

6. Outlook 加载运行时环境，这将从内容分发网络 (CDN) 服务器中为 JavaScript 库文件下载并缓存 JavaScript API，然后为 [Office](/javascript/api/office) 加载项对象的 [initialize](/javascript/api/office#Office_initialize_reason_) 事件调用事件处理程序（如果已为其分配处理程序）。 此时，它还检查 (或链接 `then()` 方法) 的任何回调是否已 (或链接) 到 `Office.onReady` 处理程序。 有关区别`Office.initialize``Office.onReady`的详细信息，请参阅[“初始化加载项](initialize-add-in.md)”。

7. 当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [初始化 Office 加载项](initialize-add-in.md)
