---
title: 在 Outlook 的 Mac 版中比较 Outlook 加载项支持
description: 了解 Mac 上的 Outlook 中的加载项支持如何与其他 Outlook 主机进行比较。
ms.date: 11/26/2019
localization_priority: Normal
ms.openlocfilehash: 337938b9bb2e8f0e9b9343841a8240e46741eed9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165937"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-hosts"></a>将 Mac 上 outlook 中的 Outlook 加载项支持与其他 Outlook 主机进行比较

您可以在 Mac 上的 Outlook 中像在其他主机（包括 Outlook 网页版、Windows、iOS 和 Android）上创建和运行 Outlook 外接程序，而无需自定义每个主机的 JavaScript。 除了下表中介绍的区域外，从外接程序到 适用于 Office 的 JavaScript API 的相同调用通常也以相同的方式进行。

有关详细信息，请参阅[部署和安装 Outlook 加载项以供测试](testing-and-tips.md)。

| 区域 | Outlook 网页版、Windows 版和移动设备 | Mac 版 Outlook |
|:-----|:-----|:-----|
| office.js 和 Office 外接程序清单架构支持的版本 | Office.js 和架构 v1.1 中的所有 API。 | Office.js 和架构 v1.1 中的所有 API。<br><br>**注意**： Mac 上的 Outlook 不支持保存会议。 在撰写模式下，无法从会议调用 `saveAsync` 方法。 若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。 |
| 定期约会系列实例 | <ul><li>可以获得主约会的项目 ID 和其他属性或定期系列约会的实例</li><li>可以使用 [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 显示定期序列的实例或主项目。</li></ul> | <ul><li>可以获得主约会的项目 ID 和其他属性，但无法获得定期系列约会的实例</li><li>可以显示定期系列的主约会。不显示项目 ID 和定期系列的实例。</li></ul> |
| 约会参与者的收件人类型 | 可以使用 [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) 标识与会者的收件人类型。 | `EmailAddressDetails.recipientType` 为约会与会者返回 `undefined`。 |
| 主机版本字符串 | [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) 返回的版本字符串的格式取决于主机的实际类型。例如：<ul><li>Windows 上的 Outlook：`15.0.4454.1002`</li><li>Outlook 网页：`15.0.918.2`</li></ul> |Outlook on Mac `Diagnostics.hostVersion`上返回的版本字符串的示例：`15.0 (140325)` |
| 项目自定义属性 | 如果网络出现故障，外接程序仍可以访问缓存的自定义属性。 | 由于 Mac 上的 Outlook 不缓存自定义属性，因此，如果网络出现故障，外接程序将无法访问它们。 |
| 附件详细信息 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) 对象中的内容类型和附件名称取决于主机的类型：<ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` 不包含任何文件名扩展名。例如，如果附件是一封主题为“RE: Summer activity”的邮件，则表示附件名称的 JSON 对象将为 `"name": "RE: Summer activity"`。</li></ul> | <ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType" "image/png"`</li><li>`AttachmentDetails.name` 始终包含一个文件名扩展名。作为邮件项目的附件包含 .eml 扩展名，约会包含 .ics 扩展名。例如，如果附件是主题为“RE: Summer activity”的电子邮件，那么表示附件名称的 JSON 对象为 `"name": "RE: Summer activity.eml"`。<p>**注意：** 如果以编程方式附加（例如通过加载项）不带扩展名的文件，`AttachmentDetails.name` 将不会在文件名中包含扩展名。</p></li></ul> |
| 表示 `dateTimeCreated` 和 `dateTimeModified` 属性中的时区的字符串 |示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` 和 `dateTimeModified` 的时间准确度 | 如果加载项使用以下代码，准确度精确到毫秒：<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 准确度精确到秒。 |

