---
title: 比较 Mac 上的 Outlook 中的 Outlook 外接程序支持
description: 了解 Mac 上的 Outlook 中的外接程序支持与其他 Outlook 客户端的对比。
ms.date: 03/19/2021
localization_priority: Normal
ms.openlocfilehash: 5a3e964b2659c6201cd22e9e55fb01d049540069
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178060"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>比较 Mac 上 Outlook 与其他 Outlook 客户端的 Outlook 外接程序支持

您可以在 Mac 版 Outlook 中像在其他客户端（包括 Outlook 网页版、Windows、iOS 和 Android 客户端）中一样创建和运行 Outlook 外接程序，而无需为每个客户端自定义 JavaScript。 从外接程序到 Office JavaScript API 的相同调用通常以相同方式工作，下表中描述的区域除外。

有关详细信息，请参阅[部署和安装 Outlook 外接程序以进行测试](testing-and-tips.md)。

有关新 UI 支持的信息，请参阅 [新 Mac UI 上的 Outlook 中的外接程序支持](#add-in-support-in-outlook-on-new-mac-ui-preview)。

| 区域 | Outlook 网页版、Windows 和移动设备 | Mac 版 Outlook |
|:-----|:-----|:-----|
| office.js 和 Office 外接程序清单架构支持的版本 | Office.js 和架构 v1.1 中的所有 API。 | Office.js 和架构 v1.1 中的所有 API。<br><br>**注意**：在 Mac 版 Outlook 中，仅内部版本 16.35.308 或更高版本支持保存会议。 否则， `saveAsync` 在撰写模式下从会议调用方法时失败。 若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。 |
| 定期约会系列实例 | <ul><li>可以获得主约会的项目 ID 和其他属性或定期系列约会的实例</li><li>可以使用 [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 显示定期序列的实例或主项目。</li></ul> | <ul><li>可以获得主约会的项目 ID 和其他属性，但无法获得定期系列约会的实例</li><li>可以显示定期系列的主约会。不显示项目 ID 和定期系列的实例。</li></ul> |
| 约会参与者的收件人类型 | 可以使用 [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) 标识与会者的收件人类型。 | `EmailAddressDetails.recipientType` 为约会与会者返回 `undefined`。 |
| 客户端应用程序的版本字符串 | [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion)返回的版本字符串的格式取决于客户端的实际类型。 例如：<ul><li>Windows 上的 Outlook： `15.0.4454.1002`</li><li>Outlook 网页： `15.0.918.2`</li></ul> |Mac 版 Outlook 上返回 `Diagnostics.hostVersion` 的版本字符串示例： `15.0 (140325)` |
| 项目自定义属性 | 如果网络出现故障，外接程序仍可以访问缓存的自定义属性。 | 由于 Mac 上的 Outlook 不缓存自定义属性，因此如果网络关闭，外接程序将无法访问它们。 |
| 附件详细信息 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)对象中的内容类型和附件名称取决于客户端的类型：<ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` 不包含任何文件名扩展名。例如，如果附件是一封主题为“RE: Summer activity”的邮件，则表示附件名称的 JSON 对象将为 `"name": "RE: Summer activity"`。</li></ul> | <ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType" "image/png"`</li><li>`AttachmentDetails.name` 始终包含一个文件名扩展名。作为邮件项目的附件包含 .eml 扩展名，约会包含 .ics 扩展名。例如，如果附件是主题为“RE: Summer activity”的电子邮件，那么表示附件名称的 JSON 对象为 `"name": "RE: Summer activity.eml"`。<p>**注意：** 如果以编程方式附加（例如通过加载项）不带扩展名的文件，`AttachmentDetails.name` 将不会在文件名中包含扩展名。</p></li></ul> |
| 表示 `dateTimeCreated` 和 `dateTimeModified` 属性中的时区的字符串 |示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` 和 `dateTimeModified` 的时间准确度 | 如果加载项使用以下代码，准确度精确到毫秒：<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 准确度精确到秒。 |

## <a name="add-in-support-in-outlook-on-new-mac-ui-preview"></a>新版 Mac UI 上的 Outlook 中的外接程序支持 (预览) 

Outlook 外接程序现在支持新的 Mac UI (预览) ，最高为要求集 1.7。 但是，尚不支持以下 **要求集** 和功能。

- API 要求集 1.8 和 1.9
- On-send
- "撰写"窗口弹出
- 共享文件夹支持
- `saveAsync` 撰写会议时

我们鼓励你在 16.38.506 版中提供的新 Mac UI 上预览 Outlook。 若要详细了解如何试用，请参阅 Outlook for [Mac - Insider Fast 内部版本发行说明](https://support.microsoft.com/office/d6347358-5613-433e-a49e-a9a0e8e0462a)。

你可以确定你位于哪个 UI 版本，如下所示。

**当前 UI**

&nbsp;&nbsp;&nbsp;&nbsp;![Mac 上的当前 UI](../images/outlook-on-mac-classic.png)

**新的 UI (预览)**

&nbsp;&nbsp;&nbsp;&nbsp;![Mac 上预览版中的新 UI](../images/outlook-on-mac-new.png)
