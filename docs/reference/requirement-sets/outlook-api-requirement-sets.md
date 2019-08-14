---
title: Outlook JavaScript API 要求集
description: ''
ms.date: 08/13/2019
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 010dd0519ff6b82f29e2ee7c3cdebb9a64106ac9
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395657"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Outlook JavaScript API 要求集

Outlook 外接程序通过在其清单中使用 Requirements 元素来声明所需要的 API 版本。Outlook 外接程序始终包括  属性设置为  和  属性设置为支持外接程序方案的 API 最低要求集的 Set 元素。

例如，下面的清单段表示 1.1 的最低要求集。

```xml
<Requirements>
  <Sets>
    <Set Name="Mailbox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

所有 Outlook API 均属于`Mailbox`[要求集](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)。`Mailbox`要求集具有不同版本，我们发布的每个新 API 集均属于较高版本的要求集。并非所有 Outlook 客户端都支持最新的 API 集，但如果某个 Outlook 客户端声明支持某个要求集，它将支持该要求集中的所有 API。

在清单中设置最低要求集版本可控制外接程序会显示在哪个 Outlook 客户端中。如果客户端不支持最低要求集，则不会加载外接程序。例如，如果指定要求集版本 1.3，则意味着外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。

> [!NOTE]
> 要在任何带编号的要求集中使用 API，应引用 CDN 上的**生产**库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js)。
>
> 要了解如何使用预览 API，请参阅本文稍后的[使用预览 API](#using-preview-apis) 部分。

## <a name="using-apis-from-later-requirement-sets"></a>使用更高版本要求集中的 API

设置要求集不会限制外接程序可使用的可用 API。 例如，如果加载项指定要求集 1.1，但它在支持版本 1.3 的 Outlook 客户端中运行，则该加载项从要求集 1.3 使用 API。

要使用较新的 API，开发人员可执行以下操作来检查特定主机是否支持相应的要求集。

```js
if (Office.context.requirements.isSetSupported('Mailbox', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

或者，开发人员可以使用标准的 JavaScript 技术检查是否存在较新 API。

```js
if (item.somePropertyOrFunction !== undefined) {
  // Use item.somePropertyOrFunction.
  item.somePropertyOrFunction;
}
```

对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。

## <a name="choosing-a-minimum-requirement-set"></a>选择最低要求集

开发人员应使用包含其方案关键 API 集的最早要求集，如果不使用该要求集，外接程序将不起作用。

## <a name="clients"></a>客户端

下列客户端支持 Outlook 外接程序。

| 客户端 | 受支持的 API 要求集 |
| --- | --- |
| Windows 版 Outlook（连接到 Office 365 订阅） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Windows 版 Outlook 2019（一次性购买） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Windows 版 Outlook 2016（一次性购买） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Windows 版 Outlook 2013（一次性购买） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Mac 版 Outlook（连接到 Office 365 订阅） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Mac 版 Outlook 2019（一次性购买） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Mac 版 Outlook 2016（一次性购买） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| iOS 版 Outlook | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Android 版 Outlook | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook 网页版（新式） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 网页版（经典） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| 连接到本地 Exchange 2019 的任何 Outlook 客户端 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| 连接到本地 Exchange 2016 的任何 Outlook 客户端 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| 连接到本地 Exchange 2013 的任何 Outlook 客户端 | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |

> [!NOTE]
> 对 Outlook 2013 中的 1.3 版本的支持已作为 [2015 年 12 月 8 日 Outlook 2013 更新 (KB3114349)](https://support.microsoft.com/kb/3114349) 的一部分添加。 对 Outlook 2013 中的 1.4 版本的支持已作为 [2016 年 9 月 13 日 Outlook 2013 更新 (KB3118280)](https://support.microsoft.com/help/3118280) 的一部分添加。 对 Outlook 2016 (MSI) 中的 1.4 版本的支持已作为 [2018 年 7 月 3 日 Office 2016 更新 (KB4022223)](https://support.microsoft.com/help/4022223) 的一部分添加。

> [!TIP]
> 可通过查看邮箱工具栏，在 Web 浏览器中区分经典和新式 Outlook。
>
> **新式**
>
> ![新式 Outlook 工具栏的部分屏幕截图](https://docs.microsoft.com/outlook/add-ins/images/outlook-on-the-web-new-toolbar.png)
>
> **经典**
>
> ![经典 Outlook 工具栏的部分屏幕截图](https://docs.microsoft.com/outlook/add-ins/images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>使用预览 API

新的 Outlook JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。 若要提供有关预览 API 的反馈，请使用在其内记录 API 的网页末尾的反馈机制。

> [!NOTE]
> 预览 API 可能会发生变更，不适合在生产环境中使用。

有关预览 API 的更多详细信息，请参阅 [Outlook API 预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。
