---
title: Outlook JavaScript API 要求集
description: 了解有关 Outlook JavaScript API 要求集的详细信息。
ms.date: 02/08/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: d3a9255ccba922ecaef5aafe8407e98d4ab2fc33
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234140"
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

所有 Outlook API 均属于`Mailbox`[要求集](../../develop/specify-office-hosts-and-api-requirements.md)。`Mailbox`要求集具有不同版本，我们发布的每个新 API 集均属于较高版本的要求集。并非所有 Outlook 客户端都支持最新的 API 集，但如果某个 Outlook 客户端声明支持某个要求集，它通常支持该需求集中的所有 API（请查看有关特定 API 或功能的文档以了解任何异常）。

在清单中设置最低要求集版本可控制外接程序会显示在哪个 Outlook 客户端中。如果客户端不支持最低要求集，则不会加载外接程序。例如，如果指定要求集版本 1.3，则意味着外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。

> [!NOTE]
> 要在任何带编号的要求集中使用 API，应引用 CDN 上的 **生产** 库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js)。
>
> 要了解如何使用预览 API，请参阅本文稍后的[使用预览 API](#using-preview-apis) 部分。

## <a name="using-apis-from-later-requirement-sets"></a>使用更高版本要求集中的 API

设置要求集不会限制外接程序可使用的可用 API。 例如，如果加载项指定要求集“Mailbox 1.1”，但它在支持版本“Mailbox 1.3”的 Outlook 客户端中运行，则该加载项从要求集“Mailbox 1.3”使用 API。

若要使用较新的 API，开发人员可执行以下操作来检查特定应用程序是否支持相应的要求集：

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

## <a name="requirement-sets-supported-by-exchange-servers-and-outlook-clients"></a>Exchange 服务器和 Outlook 客户端支持的要求集

本节将说明 Exchange 服务器和 Outlook 客户端支持的要求集范围。 有关运行 Outlook 加载项的服务器和客户端要求的详细信息，请参阅 [Outlook 加载项要求](../../outlook/add-in-requirements.md)。

> [!IMPORTANT]
> 如果目标 Exchange 服务器和 Outlook 客户端支持不同的要求集，则将受限于较低的要求集范围。 例如，如果外接程序在 Mac 上的 Outlook 2016（最高要求集：1.6）中针对 Exchange 2013（最高要求集：1.1）运行，则外接程序受限于要求集 1.1。

### <a name="exchange-server-support"></a>Exchange 服务器支持

下列服务器支持 Outlook 外接程序。

| 产品 | 主要 Exchange 版本 | 受支持的 API 要求集 |
|---|---|---|
| Exchange Online | 最新版本 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、[1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)、[1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)\* |
| 本地 Exchange | 2019 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2016 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) |
|| 2013 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md) |

> [!NOTE]
> \*若要在加载项代码中要求标识 API 设置为 1.3，请通过呼叫 `isSetSupported('IdentityAPI', '1.3')` 来检查是否其是否受到支持。 声明其在加载项清单中不受支持。 还可通过检查其不是 `undefined` 来确定该 API 是否受到支持。 有关详细信息，请参阅 [使用后续要求集中的 API](#using-apis-from-later-requirement-sets)。

### <a name="outlook-client-support"></a>Outlook 客户端支持

下列平台上的 Outlook 支持外接程序。

| 平台 | 主要 Office/Outlook 版本 | 受支持的 API 要求集 |
|---|---|---|
| Windows | Microsoft 365 订阅 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、[1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup>、[1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<sup>1</sup><br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 2019 一次性购买（零售） | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、[1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<sup>1</sup> |
|| 2019 一次性购买（批量许可） | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| 2016 一次性购买 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
|| 2013 一次性购买 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)<sup>3</sup>、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)<sup>3</sup> |
| Mac | 当前 UI<br>（关联至 Microsoft 365 订阅） | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、[1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 新建 UI（预览版）<sup>4</sup><br>（关联至 Microsoft 365 订阅） | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md) |
|| 2019 一次性购买 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
|| 2016 一次性购买 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |
| iOS | Microsoft 365 订阅 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Android | Microsoft 365 订阅 | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)<sup>5</sup> |
| Web 浏览器 | 连接到的新式 Outlook UI<br>Exchange Online：Microsoft 365 订阅、Outlook.com | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md)、[1.7](../objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md)、[1.8](../objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)、[1.9](../objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)<br>[IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)<sup>2</sup> |
|| 连接到的经典 Outlook UI<br>本地 Exchange | [1.1](../objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md)、[1.2](../objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md)、[1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)、[1.4](../objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md)、[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)、[1.6](../objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md) |

> [!NOTE]
> <sup>1</sup>自版本 1910（内部版本 12130.20272）起，包含 Microsoft 365 订阅或零售一次性购买的 Outlook on Windows 开始支持 **1.8**。 版本 2008（内部版本 13127.20296）在包含 Microsoft 365 订阅的 Windows 版 Outlook 中支持 **1.9** 版本。 如需了解你的版本的更多详情，请参阅 [Office 2019](/officeupdates/update-history-office-2019) 或 [Microsoft 365](/officeupdates/update-history-office365-proplus-by-date) 的更新历史记录页，以及如何[查找 Office 客户端版本和更新通道](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)。
>
> <sup>2</sup> 需要加载项代码中的 Identity API set 1.3，请通过呼叫 `isSetSupported('IdentityAPI', '1.3')` 检查其是否受到支持。 声明其在加载项清单中不受支持。 还可通过检查其不是 `undefined` 来确定该 API 是否受到支持。 有关详细信息，请参阅 [使用后续要求集中的 API](#using-apis-from-later-requirement-sets)。
>
> <sup>3</sup> [2015 年 12 月 8 日 Outlook 2013 更新 (KB3114349)](https://support.microsoft.com/kb/3114349) 开始在 Outlook 2013 中支持 1.3 版本。 对 Outlook 2013 中的 1.4 版本的支持已作为 [2016 年 9 月 13 日 Outlook 2013 更新 (KB3118280)](https://support.microsoft.com/help/3118280) 的一部分添加。 对 Outlook 2016（一次性购买）中的 1.4 版本的支持已作为 [2018 年 7 月 3 日 Office 2016 更新 (KB4022223)](https://support.microsoft.com/help/4022223) 的一部分添加。
>
> <sup>4</sup> Outlook 版本 16.38.506 已提供对全新 Mac UI（预览版）的支持。 有关详细信息，请参阅 [全新 Mac UI 上 Outlook 中的加载项支持](../../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview) 部分。
>
> <sup>5</sup> 目前，设计和实现移动客户端的加载项时有其他注意事项。 例如，只支持“邮件阅读”模式。 有关更多详细信息，请参阅[为 Outlook Mobile 添加加载项命令支持时的代码注意事项](../../outlook/add-mobile-support.md#code-considerations)。

> [!TIP]
> 可通过查看邮箱工具栏，在 Web 浏览器中区分经典和新式 Outlook。
>
> **新式**
>
> ![新式 Outlook 工具栏的部分屏幕截图](../../images/outlook-on-the-web-new-toolbar.png)
>
> **经典**
>
> ![经典 Outlook 工具栏的部分屏幕截图](../../images/outlook-on-the-web-classic-toolbar.png)

## <a name="using-preview-apis"></a>使用预览 API

新的 Outlook JavaScript API 首先在“预览版”中引入，在进行充分测试并获得用户反馈后，它将成为编号的特定要求集的一部分。 若要提供有关预览 API 的反馈，请使用在其内记录 API 的网页末尾的反馈机制。

> [!NOTE]
> 预览 API 可能会发生变更，不适合在生产环境中使用。

有关预览 API 的更多详细信息，请参阅 [Outlook API 预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。
