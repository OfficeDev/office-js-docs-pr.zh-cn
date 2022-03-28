---
title: 获取和设置 Internet 标头
description: 如何获取和设置加载项中邮件的Outlook标头。
ms.date: 04/28/2020
ms.localizationpriority: medium
ms.openlocfilehash: ddbb555f8901e1b244fb3e30682d73c21928963e
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483998"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>在加载项中获取和设置Outlook邮件头

## <a name="background"></a>背景

在外接程序Outlook中的一个常见要求是，在不同级别存储与外接程序关联的自定义属性。 目前，自定义属性存储在项目或邮箱级别。

- 项目级别 - 对于应用于特定项目的属性，请使用 [CustomProperties](/javascript/api/outlook/office.customproperties) 对象。 例如，存储与发送电子邮件的人相关联的客户代码。
- 邮箱级别 - 对于适用于用户邮箱中所有邮件项目的属性，请使用 [RoamingSettings](/javascript/api/outlook/office.roamingsettings) 对象。 例如，存储用户的首选项以以特定刻度显示温度。

这两种类型的属性在项目从服务器Exchange后不会保留，因此电子邮件收件人无法获取在项目上设置的任何属性。 因此，开发人员无法访问这些设置或其他 MIME 属性，从而启用更好的读取方案。

虽然可以通过 EWS 请求设置 Internet 标头，但在某些情况下，EWS 请求不起作用。 例如，在桌面Outlook撰写模式下，项目 ID `saveAsync`  不会在缓存模式下同步。

> [!TIP]
> 请参阅 [Get and set add-in metadata for an Outlook in to](metadata-for-an-outlook-add-in.md) learn more about using these options.

## <a name="purpose-of-the-internet-headers-api"></a>Internet 标头 API 的用途

要求 [集 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) 中引入了 Internet 标头 API，开发人员可以：

- 标记电子邮件在离开所有客户端后Exchange的信息。
- 在邮件阅读方案中，阅读电子邮件离开Exchange保留的电子邮件的信息。
- 访问电子邮件的整个 MIME 标头。

![Internet 标头关系图。 文本：用户 1 发送电子邮件。 在用户撰写电子邮件时，外接程序管理自定义 Internet 标头。 用户 2 接收电子邮件。 外接程序从收到的电子邮件获取 Internet 标头，然后分析和使用自定义标头。](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>在撰写邮件时设置 Internet 标头

请尝试使用 [item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) 属性管理在撰写模式下在当前邮件上放置的自定义 Internet 标头。

### <a name="set-get-and-remove-custom-headers-example"></a>设置、获取和删除自定义标头示例

以下示例演示如何设置、获取和删除自定义标头。

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

## <a name="get-internet-headers-while-reading-a-message"></a>在阅读邮件时获取 Internet 标头

尝试调用 [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)) ，以在阅读模式下获取当前邮件上的 Internet 标头。

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>从当前 MIME 头获取发件人首选项示例

基于上一部分的示例，以下代码显示如何从当前电子邮件的 MIME 邮件头获取发件人的首选项。

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
> 此示例适用于简单情况。 有关更复杂的信息检索 (如 [RFC 2822](https://tools.ietf.org/html/rfc2822)) 中所述的多实例标头或折叠值，请尝试使用适当的 MIME 分析库。

## <a name="recommended-practices"></a>建议的做法

目前，Internet 邮件头是用户邮箱上的有限资源。 当配额用尽时，你无法在此邮箱上再创建一个 Internet 标头，这可能会导致依赖此功能的客户端发生意外行为。

在外接程序中创建 Internet 标头时，请应用以下准则。

- 创建所需的最小标头数。
- 命名标头，以便以后可以重复使用和更新其值。 因此，避免以可变 (命名标头，例如，根据用户输入、时间戳等) 。

## <a name="see-also"></a>另请参阅

- [获取和设置 Outlook 加载项的元数据](metadata-for-an-outlook-add-in.md)
