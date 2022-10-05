---
title: Outlook 加载项 API
description: 了解如何引用 Outlook 加载项 API 并声明 Outlook 加载项中的权限。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 69043646add5e32502efb0d2a5b1259667e564d9
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467074"
---
# <a name="outlook-add-in-apis"></a>Outlook 外接程序 API

要将 API 用于您的 Outlook 外接程序，您必须指定 Office.js 库的位置、要求集、架构和权限。 你将主要使用通过 [邮箱](#mailbox-object) 对象公开的 Office JavaScript API。

## <a name="officejs-library"></a>Office.js 库

若要与 [Outlook 外接程序 API 交互](/javascript/api/outlook)，需要在Office.js中使用 JavaScript API。 库的内容分发网络 (CDN) 。`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js` 提交到 AppSource 的加载项必须按此 CDN 引用 Office.js，它们不能使用本地引用。

在实现加载项 UI 的网页（.html、.aspx 或 .php 文件）的 `<head>` 标记的 `<script>` 标记中引用 CDN。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.

> [!IMPORTANT]
> 为任何 Office 客户端应用程序开发加载项时，请从页面的内部 `<head>` 引用 Office JavaScript API。 这样可确保 API 先于所有正文元素完全初始化。

## <a name="requirement-sets"></a>要求集

所有 Outlook API 都属于 [邮箱要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)。 `Mailbox` 要求集具有多个版本，并且我们发布的每组新的 API 都属于更高版本的要求集。 并非所有的 Outlook 客户端在发布时都将支持最新的 API 集，但如果 Outlook 客户端声明支持要求集，它将支持该要求集中的所有 API。

To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in will not show up in any Outlook client that doesn't support a minimum version of 1.3.

Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.

要检查大于清单中所指定要求集的要求集中任何 API 的可用性，可以使用标准 JavaScrip：

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> 对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。

指定支持你的方案的关键 API 集的最低要求集，如果缺少该要求集，加载项的功能将无法工作。 在清单中指定要求集。 标记因所使用的清单而异。 

- **XML 清单**：使用该 **\<Requirements\>** 元素。 请注意， **\<Methods\>** Outlook 外接程序不支持其子元素 **\<Requirements\>** ，因此不能声明对特定方法的支持。
- **Teams 清单 (预览)**：使用“extensions.capabilities”属性。 

有关详细信息，请参阅 [Outlook 加载项清单](manifests.md)和 [了解 Outlook API 要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)。

## <a name="permissions"></a>权限

外接程序需要相应的权限才能使用所需的 API。 一般情况下，应该指定加载项所需的最小权限。

有四个级别的权限; **受限**、 **读取项目**、 **读/写项目** 和 **读/写邮箱**。 有关更多详细信息。 有关详细信息，请参阅[了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)。

## <a name="mailbox-object"></a>Mailbox 对象

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>另请参阅

- [Outlook 加载项清单](manifests.md)
- [了解 Outlook API 要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [了解 Outlook 加载项权限](understanding-outlook-add-in-permissions.md)。
- [Office 加载项的隐私和安全](../concepts/privacy-and-security.md)
