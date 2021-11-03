---
title: Outlook外接程序 API 预览要求集
description: 当前处于预览阶段的功能和 API Outlook外接程序。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2d1efa2b2dca5a88a56fb5f54a84b790e08745ec
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681646"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook外接程序 API 预览要求集

Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!IMPORTANT]
> 本文档适用于 **预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> 你或许能够通过在 Outlook 网页版 租户上配置目标版本来预览 Microsoft 365[功能](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。 此页面中会针对适用的功能说明"配置预览访问"。
>
> 对于其他功能，您可能能够通过完成和提交此表单Outlook 网页版请求对预览位Microsoft 365帐户进行[预览](https://aka.ms/OWAPreview)。 这些功能中会指出"请求预览访问"。

预览要求集包含要求集 [1.11 的所有功能](../requirement-set-1.11/outlook-requirement-set-1.11.md)。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>对受信息权限管理中心 IRM 保护的项目 (加载项) 

现在可以在受 IRM 保护的项目上激活外接程序。 若要启用此功能，租户管理员需要在租户中设置"允许以编程方式访问"自定义策略选项，以 `OBJMODEL` 启用Office。  有关详细信息 [，请参阅使用](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) 权限和说明。

**适用于**：Outlook（Windows版本 13229.10000） (连接到 Microsoft 365 订阅) 

<br>

---

---

### <a name="additional-calendar-properties"></a>其他日历属性

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

添加了一个新对象，该对象代表撰写模式下约会的全天事件属性。

**在**：Outlook Windows (订阅Microsoft 365上) 

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

添加了一个新对象，该对象表示撰写模式下约会的敏感度。

**在**：Outlook Windows (订阅Microsoft 365上) 

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

添加了一个新属性，它表示约会是全天事件。

**在**：Outlook Windows (订阅Microsoft 365上) 

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

新增了一个表示约会敏感度的属性。

**在**：Outlook Windows (订阅Microsoft 365上) 

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office。MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

新增了表示约会 `AppointmentSensitivityType` 可用的敏感度选项的枚举。

**在**：Outlook Windows (订阅Microsoft 365上) 

<br>

---

---

### <a name="delay-delivery-time"></a>延迟传递时间

#### <a name="officecontextmailboxitemdelaydeliverytime"></a>[Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties)

添加了一个新属性，该属性返回一个对象，该对象允许您在撰写模式下管理邮件的传递日期和时间。

**在**：Outlook Windows (订阅Microsoft 365上) 

#### <a name="officedelaydeliverytime"></a>[Office。DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true)

新增了一个对象，允许您在撰写模式下管理邮件的传递日期和时间。

**在**：Outlook Windows (订阅Microsoft 365上) 

<br>

---

---

### <a name="event-based-activation"></a>基于事件的激活

此功能在要求集 [1.10 中发布](../requirement-set-1.10/outlook-requirement-set-1.10.md)。 不过，其他事件现在预览版中可用。 若要了解详情，请参阅 [支持的事件](../../../outlook/autolaunch.md#supported-events)。

**在**：Outlook Windows (订阅Microsoft 365上) 

<br>

---

---

### <a name="integration-with-actionable-messages"></a>与可操作邮件集成

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。

**适用于**：Outlook Windows (连接到 Microsoft 365 订阅) ，Outlook 网页版 (新式) 

<br>

---

---

### <a name="office-theme"></a>Office 主题

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#officeTheme)

增加了获取 Office 主题的功能。

**在**：Outlook Windows (订阅Microsoft 365上) 

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。

**在**：Outlook Windows (订阅Microsoft 365上) 

<br>

---

---

### <a name="shared-mailboxes"></a>共享邮箱

共享文件夹的功能支持 (，即委派访问权限) 要求集 [1.8 中发布](../requirement-set-1.8/outlook-requirement-set-1.8.md)。 但是，现在预览版中提供了对共享邮箱的支持。 要了解详细信息，请参阅 [启用共享文件夹和共享邮箱方案](../../../outlook/delegate-access.md)。

**适用于**：Outlook Windows (连接到 Microsoft 365 订阅) 、Outlook 网页版 (新式) 、Outlook Mac 上

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
