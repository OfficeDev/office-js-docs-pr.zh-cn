---
title: 比较 Mac 上Outlook中的Outlook加载项支持
description: 了解 Mac Outlook中的外接程序支持与其他Outlook客户端的比较。
ms.date: 06/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: 36a10f0454bebf3f069464277c7eb2a8a18f42b7
ms.sourcegitcommit: 2eeb0423a793b3a6db8a665d9ae6bcb10e867be3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/10/2022
ms.locfileid: "66019603"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>将 Mac 上Outlook中的Outlook外接程序支持与其他Outlook客户端进行比较

在 Mac 上Outlook，可以像在其他客户端（包括Outlook 网页版、Windows、iOS和Android）中一样创建和运行Outlook加载项，而无需为每个客户端自定义 JavaScript。 从外接程序到 Office JavaScript API 的相同调用通常以相同方式工作，但下表中所述的区域除外。

有关详细信息，请参阅[部署和安装 Outlook 外接程序以进行测试](testing-and-tips.md)。

有关新的 UI 支持的信息，请参阅[新 Mac UI Outlook中的加载项支持](#add-in-support-in-outlook-on-new-mac-ui)。

| 领域 | Outlook 网页版、Windows 和移动设备 | Mac 版 Outlook |
|:-----|:-----|:-----|
| office.js 和 Office 外接程序清单架构支持的版本 | Office.js 和架构 v1.1 中的所有 API。 | Office.js 和架构 v1.1 中的所有 API。<br><br>**注意**：在 Mac 上的Outlook中，仅内部版本 16.35.308 或更高版本支持保存会议。 否则，方法在 `saveAsync` 撰写模式下从会议调用时失败。 若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。 |
| 定期约会系列实例 | <ul><li>可以获得主约会的项目 ID 和其他属性或定期系列约会的实例</li><li>可以使用 [mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 显示定期序列的实例或主项目。</li></ul> | <ul><li>可以获得主约会的项目 ID 和其他属性，但无法获得定期系列约会的实例</li><li>可以显示定期系列的主约会。不显示项目 ID 和定期系列的实例。</li></ul> |
| 约会参与者的收件人类型 | 可以使用 [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) 标识与会者的收件人类型。 | `EmailAddressDetails.recipientType` 为约会与会者返回 `undefined`。 |
| 客户端应用程序的版本字符串 | [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) 返回的版本字符串的格式取决于客户端的实际类型。 例如：<ul><li>Windows上的Outlook：`15.0.4454.1002`</li><li>Outlook 网页版：`15.0.918.2`</li></ul> |Mac 上Outlook返回的`Diagnostics.hostVersion`版本字符串示例：`15.0 (140325)` |
| 项目自定义属性 | 如果网络出现故障，外接程序仍可以访问缓存的自定义属性。 | 由于 Mac 上的Outlook不会缓存自定义属性，因此如果网络关闭，加载项将无法访问它们。 |
| 附件详细信息 | [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) 对象中的内容类型和附件名称取决于客户端的类型：<ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType": "image/x-png"`。 </li><li>`AttachmentDetails.name` 不包含任何文件名扩展名。例如，如果附件是一封主题为“RE: Summer activity”的邮件，则表示附件名称的 JSON 对象将为 `"name": "RE: Summer activity"`。</li></ul> | <ul><li>`AttachmentDetails.contentType` 的 JSON 示例：`"contentType" "image/png"`</li><li>`AttachmentDetails.name` 始终包含一个文件名扩展名。作为邮件项目的附件包含 .eml 扩展名，约会包含 .ics 扩展名。例如，如果附件是主题为“RE: Summer activity”的电子邮件，那么表示附件名称的 JSON 对象为 `"name": "RE: Summer activity.eml"`。<p>**注意：** 如果以编程方式附加（例如通过加载项）不带扩展名的文件，`AttachmentDetails.name` 将不会在文件名中包含扩展名。</p></li></ul> |
| 表示 `dateTimeCreated` 和 `dateTimeModified` 属性中的时区的字符串 |示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | 示例：`Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| `dateTimeCreated` 和 `dateTimeModified` 的时间准确度 | 如果外接程序使用以下代码，准确度精确到毫秒。<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| 准确度精确到秒。 |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>新 Mac UI 上Outlook中的加载项支持

Outlook新版 Mac UI (从 Outlook 版本 16.38.506) （至要求设置 1.10）支持加载项。 但是，尚 **不** 支持以下要求集和功能。

- API 要求集 1.11

若要详细了解新的 Mac UI，请参阅[新Outlook for Mac](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439)。

可以确定所使用的 UI 版本，如下所示：

**经典 UI**

![Mac 上的经典 UI。](../images/outlook-on-mac-classic.png)

**新建 UI**

![Mac 上的新 UI。](../images/outlook-on-mac-new.png)
