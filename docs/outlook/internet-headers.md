---
title: 获取和设置 Internet 标头
description: 如何在 Outlook 加载项中获取和设置邮件上的 Internet 标头。
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8e4af70b24a96b8d00acc7ea4101acf53e2b71
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616026"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>在 Outlook 外接程序中获取和设置邮件上的 Internet 标头

## <a name="background"></a>背景

Outlook 外接程序开发中的一个常见要求是在不同级别存储与加载项关联的自定义属性。 目前，自定义属性存储在项目或邮箱级别。

- 项级别 - 对于应用于特定项的属性，请使用 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象。 例如，存储与发送电子邮件的人员关联的客户代码。
- 邮箱级别 - 对于适用于用户邮箱中所有邮件项的属性，请使用 [RoamingSettings](/javascript/api/outlook/office.roamingsettings) 对象。 例如，存储用户的首选项以显示特定比例的温度。

项目离开 Exchange 服务器后不会保留这两种类型的属性，因此电子邮件收件人无法获取对该项设置的任何属性。 因此，开发人员无法访问这些设置或其他多用途 Internet 邮件扩展 (MIME) 属性，以实现更好的读取方案。

虽然可以通过 Exchange Web 服务 (EWS) 请求来设置 Internet 标头，但在某些情况下，发出 EWS 请求将不起作用。 例如，在 Outlook 桌面上的 Compose 模式下，项目 ID 不会在缓存模式下 `saveAsync` 同步。

> [!TIP]
> 若要详细了解如何使用这些选项，请参阅 [获取和设置 Outlook 外接程序的加载项元](metadata-for-an-outlook-add-in.md)数据。

## <a name="purpose-of-the-internet-headers-api"></a>Internet 标头 API 的用途

在 [要求集 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) 中引入，Internet 标头 API 使开发人员能够：

- 在所有客户端离开 Exchange 后保留的电子邮件的标记信息。
- 读取电子邮件在邮件读取方案中所有客户端离开 Exchange 后保留的电子邮件的信息。
- 访问电子邮件的整个 MIME 标头。

![Internet 标头图。 文本：用户 1 发送电子邮件。 加载项在用户撰写电子邮件时管理自定义 Internet 标头。 用户 2 接收电子邮件。 外接程序从收到的电子邮件中获取 Internet 标头，然后分析并使用自定义标头。](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>撰写消息时设置 Internet 标头

使用 [item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) 属性管理在 Compose 模式下当前消息上放置的自定义 Internet 标头。

### <a name="set-get-and-remove-custom-internet-headers-example"></a>设置、获取和删除自定义 Internet 标头示例

以下示例演示如何设置、获取和删除自定义 Internet 标头。

```js
// Set custom internet headers.
function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "x-preferred-fruit": "orange", "x-preferred-vegetable": "broccoli", "x-best-vegetable": "spinach" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["x-preferred-fruit", "x-preferred-vegetable", "x-best-vegetable", "x-nonexistent-header"],
    getCallback
  );
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}

// Remove custom internet headers.
function removeSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(
    ["x-best-vegetable", "x-nonexistent-header"],
    removeCallback);
}

function removeCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully removed selected headers");
  } else {
    console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
  }
}

setCustomHeaders();
getSelectedCustomHeaders();
removeSelectedCustomHeaders();
getSelectedCustomHeaders();

/* Sample output:
Successfully set headers
Selected headers: {"x-best-vegetable":"spinach","x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
Successfully removed selected headers
Selected headers: {"x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
*/
```

## <a name="get-internet-headers-while-reading-a-message"></a>读取消息时获取 Internet 标头

调用 [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)) 以在读取模式下获取当前消息上的 Internet 标头。

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>从当前 MIME 标头示例获取发件人首选项

在上一部分的示例的基础上，以下代码演示如何从当前电子邮件的 MIME 标头获取发件人的首选项。

```js
Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Sender's preferred fruit: " + asyncResult.value.match(/x-preferred-fruit:.*/gim)[0].slice(19));
    console.log("Sender's preferred vegetable: " + asyncResult.value.match(/x-preferred-vegetable:.*/gim)[0].slice(23));
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}

/* Sample output:
Sender's preferred fruit: orange
Sender's preferred vegetable: broccoli
*/
```

> [!IMPORTANT]
> 此示例适用于简单事例。 若要更复杂的信息检索 (例如，如 [RFC 2822](https://tools.ietf.org/html/rfc2822)) 中所述的多实例标头或折叠值，请尝试使用适当的 MIME 分析库。

## <a name="recommended-practices"></a>建议的做法

目前，Internet 标头是用户邮箱上的有限资源。 当配额耗尽时，无法在该邮箱上再创建任何 Internet 标头，这可能会导致依赖此功能的客户端出现意外行为。

在外接程序中创建 Internet 标头时，请应用以下准则。

- 创建所需的最小标头数。 标头配额基于应用于消息的标头的总大小。 在Exchange Online中，标头限制上限为 256 KB，而在 Exchange 本地环境中，限制由组织的管理员确定。 有关标头限制的详细信息，请[参阅Exchange Online消息限制](/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits)和[Exchange Server消息限制](/exchange/mail-flow/message-size-limits)。
- 命名标头，以便稍后可以重复使用和更新其值。 因此，请避免以可变的方式命名标头 (例如，基于用户输入、时间戳等) 。

## <a name="see-also"></a>另请参阅

- [获取和设置 Outlook 加载项的元数据](metadata-for-an-outlook-add-in.md)
