---
title: Outlook JavaScript API 要求集
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: 3d2b17de4e1bc8510b06901b4cfd1949d9490564
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433024"
---
# <a name="outlook-javascript-api-requirement-sets"></a>Outlook JavaScript API 要求集

Outlook 加载项通过在其[清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)中使用 [Requirements](/office/dev/add-ins/reference/manifest/requirements) 元素来声明所需要的 API 版本。 Outlook 加载项始终包括 [Set](/office/dev/add-ins/reference/manifest/set) 元素，其中 `Name` 属性设置为 `Mailbox` 且 `MinVersion` 属性设置为支持加载项方案的 API 最低要求集。

例如，下面的清单段表示 1.1 的最低要求集：

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

所有 Outlook API 都属于 `Mailbox` [要求集](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)。 `Mailbox` 要求集具有多个版本，并且我们发布的每组新的 API 都属于更高版本的要求集。 并非所有的 Outlook 客户端都支持最新的 API 集，但如果 Outlook 客户端声明支持要求集，则它支持该要求集中的所有 API。

在清单中设置最低要求集版本可控制外接程序会显示在哪个 Outlook 客户端中。如果客户端不支持最低要求集，则不会加载外接程序。例如，如果指定要求集版本 1.3，则意味着外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。

## <a name="using-apis-from-later-requirement-sets"></a>使用更高版本要求集中的 API

设置要求集不会限制外接程序可使用的可用 API。 例如，如果外接程序指定要求集 1.1，但它在支持 1.3 的 Outlook 客户端中运行，则该外接程序可以使用要求集 1.3 中的 API。

要使用较新的 API，开发人员只需使用标准 JavaScript 技术来检查是否存在新 API：

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。

## <a name="choosing-a-minimum-requirement-set"></a>选择最低要求集

开发人员应使用包含其方案关键 API 集的最早要求集，如果不使用该要求集，外接程序将不起作用。

## <a name="clients"></a>客户端

下列客户端支持 Outlook 外接程序。

| 客户端 | 受支持的 API 要求集 |
| --- | --- |
| Outlook 2019 for Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 for Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016（即点即用）for Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6)、[1.7](/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 (MSI) for Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2016 for Mac | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2013 for Windows | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook for iPhone | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook for Android | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook 网页版（Office 365 和 Outlook.com） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5)、[1.6](/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook Web App（本地 Exchange 2013） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |
| Outlook Web App（本地 Exchange 2016） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |
| Outlook Web App（本地 Exchange 2019） | [1.1](/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1)、[1.2](/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2)、[1.3](/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3)、[1.4](/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4)、[1.5](/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |

> [!NOTE]
> 对 Outlook 2013 中的 1.3 版本的支持已作为 [2015 年 12 月 8 日 Outlook 2013 更新 (KB3114349)](https://support.microsoft.com/kb/3114349) 的一部分添加。 对 Outlook 2013 中的 1.4 版本的支持已作为 [2016 年 9 月 13 日 Outlook 2013 更新 (KB3118280)](https://support.microsoft.com/help/3118280) 的一部分添加。 对 Outlook 2016 (MSI) 中的 1.4 版本的支持已作为 [2018 年 7 月 3 日 Office 2016 更新 (KB4022223)](https://support.microsoft.com/help/4022223) 的一部分添加。
