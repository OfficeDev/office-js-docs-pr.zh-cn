---
title: Outlook 加载项 API
description: 了解如何引用 Outlook 加载项 API 并声明 Outlook 加载项中的权限。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: bd7f3b5a1b52ec3ca7a48ae7a2d467c6cd30f1e4
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325467"
---
# <a name="outlook-add-in-apis"></a>Outlook 外接程序 API

要将 API 用于您的 Outlook 外接程序，您必须指定 Office.js 库的位置、要求集、架构和权限。 你将主要使用通过[邮箱](#mailbox-object)对象公开的 Office JavaScript api。

## <a name="officejs-library"></a>Office.js 库

若要与 Outlook 加载项 API 进行交互，需要在 Office.js 中使用 JavaScript API。 库的 CDN 为 `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`。 提交到 AppSource 的加载项必须按此 CDN 引用 Office.js，它们不能使用本地引用。

在实现加载项 UI 的网页（.html、.aspx 或 .php 文件）的 `<head>` 标记的 `<script>` 标记中引用 CDN。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
添加 API 时，Office.js 的 URL 将保持不变。仅当我们打破现有的 API 行为时，才会更改 URL 中的版本。

> [!IMPORTANT]
> 为任何 Office 主机应用程序开发外接程序时，请从页面的`<head>`部分中引用 OFFICE JavaScript API。 这样可确保 API 先于所有正文元素完全初始化。 Office 主机要求外接程序在激活 5 秒钟内进行初始化。 超过此阈值会导致声明的外接程序无响应，并且会向用户显示错误消息。

## <a name="requirement-sets"></a>要求集

所有 Outlook API 都属于 `Mailbox` 要求集。 `Mailbox` 要求集具有多个版本，并且我们发布的每组新的 API 都属于更高版本的要求集。 并非所有的 Outlook 客户端在发布时都将支持最新的 API 集，但如果 Outlook 客户端声明支持要求集，它将支持该要求集中的所有 API。

若要控制外接程序在哪些 Outlook 客户端中显示，请在清单中指定最低要求集版本。例如，如果你指定要求集版本 1.3，则外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。

指定要求集不会将外接程序限定于该版本中的 API。如果外接程序指定要求集 v1.1，却在支持 v1.3 的 Outlook 客户端中运行，该外接程序仍可以使用 v1.3 API。要求集仅控制外接程序在哪些 Outlook 客户端中显示。

要检查大于清单中所指定要求集的要求集中任何 API 的可用性，可以使用标准 JavaScrip：

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> 对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。

指定支持你的方案的关键 API 集的最低要求集，如果缺少该要求集，加载项的功能将无法工作。 指定 `<Requirements>` 元素的清单中的要求集。 有关更多信息，请参阅 [Outlook 加载项清单](manifests.md)和[了解 Outlook API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md)。

`<Methods>` 元素不适用于 Outlook 加载项，因此，无法声明对特定方法的支持。

## <a name="permissions"></a>权限

外接程序需要相应的权限才能使用所需的 API。有四个级别的权限。有关详细信息，请参阅[了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)。

<br/>

|权限级别|说明|
|:-----|:-----|
| **受限** | 允许使用实体，但不允许使用正则表达式。 |
| **读取项** | 除了**受限**所允许的权限，它还允许：<ul><li>正则表达式</li><li>Outlook 外接程序 API 读取访问</li><li>获取项属性和回调令牌</li></ul> |
| **读/写** | 除了**读取项**所允许的权限，它还允许：<ul><li>Outlook 加载项 API 的完全访问权限，但不包括 `makeEwsRequestAsync`</li><li>设置项属性</li></ul> |
| **读/写邮箱** | 除了**读/写**所允许的权限，它还允许：<ul><li>创建、读取、写入项和文件夹</li><li>发送项目</li><li>调用 [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)</li></ul> |

一般情况下，应该指定加载项所需的最小权限。 权限在清单的 `<Permissions>` 元素中声明。 有关更多信息，请参阅 [Outlook 加载项清单](manifests.md)。 有关安全问题的信息，请参阅[Office 外接程序的隐私和安全性](../concepts/privacy-and-security.md)。

## <a name="mailbox-object"></a>Mailbox 对象

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>另请参阅

- [Outlook 加载项清单](manifests.md)
- [了解 Outlook API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [Office 加载项的隐私和安全](../concepts/privacy-and-security.md)
