---
title: 在 Outlook 的 Mac 版中比较 Outlook 加载项支持
description: 了解 Mac 上的 Outlook 中的加载项支持如何与其他 Outlook 客户端进行比较。
ms.date: 10/20/2020
localization_priority: Normal
ms.openlocfilehash: f63c27611115e7bc262b43118ec749b0cbbe8416
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626433"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>将 Mac 上 outlook 加载项支持与其他 Outlook 客户端进行比较

您可以像在其他客户端（包括 Outlook 网页版、Windows、iOS 和 Android）中一样，在 Mac 上以相同的方式创建和运行 Outlook 加载项，而无需自定义每个客户端的 JavaScript。 除了下表中所述的区域，从外接程序到 Office JavaScript API 的相同调用通常的工作方式相同。

有关详细信息，请参阅[部署和安装 Outlook 外接程序以进行测试](testing-and-tips.md)。

有关 Mac 上新的 UI 支持的信息，请参阅 [新建 Outlook On mac](#new-outlook-on-mac-preview)。

| 区域 | Outlook 网页版、Windows 版和移动设备 | Mac 版 Outlook |
|:-----|:-----|:-----|
| office.js 和 Office 外接程序清单架构支持的版本 | Office.js 和架构 v1.1 中的所有 API。 | Office.js 和架构 v1.1 中的所有 API。<br><br>**注意**：在 Mac 上的 Outlook 中，仅生成16.35.308 或更高版本支持保存会议。 否则，在 `saveAsync` 撰写模式下从会议中调用时，此方法将失败。 若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。 |
| 定期约会系列实例 | <ul><li>可以获得主约会的项目 ID 和其他属性或定期系列约会的实例</li><li>可以使用 [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 显示定期序列的实例或主项目。</li></ul> | <ul><li>可以获得主约会的项目 ID 和其他属性，但无法获得定期系列约会的实例</li><li>可以显示定期系列的主约会。不显示项目 ID 和定期系列的实例。</li></ul> |
| 约会参与者的收件人类型 | 可以使用 [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) 标识与会者的收件人类型。 | `EmailAddressDetails.recipientType` 为约会与会者返回 `undefined`。 |
| 客户端应用程序的版本字符串 | 由 [diagnostics.hostversion](/javascript/api/outlook/office.diagnostics#hostversion) 返回的版本字符串的格式取决于客户端的实际类型。 例如：<ul><li>Windows 上的 Outlook： `15.0.4454.1002`</li><li>Outlook 网页： `15.0.918.2`</li></ul> |Outlook on Mac 上返回的版本字符串的示例 `Diagnostics.hostVersion` ： `15.0 (140325)` |
| 项目自定义属性 | 如果网络出现故障，外接程序仍可以访问缓存的自定义属性。 | 由于 Mac 上的 Outlook 不缓存自定义属性，因此，如果网络出现故障，外接程序将无法访问它们。 |
| 附件详细信息 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)对象中的内容类型和附件名称取决于客户端的类型：<ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` 不包含任何文件名扩展名。例如，如果附件是一封主题为“RE: Summer activity”的邮件，则表示附件名称的 JSON 对象将为 `"name": "RE: Summer activity"`。</li></ul> | <ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType" "image/png"`</li><li>`AttachmentDetails.name` 始终包含一个文件名扩展名。作为邮件项目的附件包含 .eml 扩展名，约会包含 .ics 扩展名。例如，如果附件是主题为“RE: Summer activity”的电子邮件，那么表示附件名称的 JSON 对象为 `"name": "RE: Summer activity.eml"`。<p>**注意：** 如果以编程方式附加（例如通过加载项）不带扩展名的文件，`AttachmentDetails.name` 将不会在文件名中包含扩展名。</p></li></ul> |
| 表示 `dateTimeCreated` 和 `dateTimeModified` 属性中的时区的字符串 |示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` 和 `dateTimeModified` 的时间准确度 | 如果加载项使用以下代码，准确度精确到毫秒：<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 准确度精确到秒。 |

## <a name="new-outlook-on-mac-preview"></a>新的 Outlook for Mac (preview) 

Outlook 外接程序现在在新的 Mac UI （最高为要求集1.7）中受支持。 但是，尚 **不** 支持以下要求集和功能。

1. API 要求集1.8 和1。9
1. 上下文加载项
1. 发送时
1. 撰写窗口弹出窗口
1. 共享文件夹支持
1. `saveAsync` 撰写会议时

我们鼓励您预览新的 Outlook on Mac，可从版本16.38.506。 若要了解有关如何试用的详细信息，请参阅 [适用于内部版本快速生成的 Outlook For Mac 发行说明](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a)。

您可以确定您所处的 UI 版本，如下所示。

**当前 UI**

&nbsp;&nbsp;&nbsp;&nbsp;![Mac 上的当前 UI](../images/outlook-on-mac-classic.png)

** (预览的新 UI) **

&nbsp;&nbsp;&nbsp;&nbsp;![在 Mac 上预览中的新 UI](../images/outlook-on-mac-new.png)
