---
title: Office.context.mailbox.diagnostics - 要求集 1.2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: f2e613816884a5c1c00e5b96565d378434747e8e
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231268"
---
# <a name="diagnostics"></a>diagnostics

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

将诊断信息提供给 Outlook 外接程序。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

### <a name="members"></a>Members

#### <a name="hostname-string"></a>hostName: String

获取表示主机应用程序的名称的字符串。

可以是下列值之一的字符串：`Outlook`、`OutlookIOS` 或 `OutlookWebApp`。

##### <a name="type"></a>类型

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

#### <a name="hostversion-string"></a>Diagnostics.hostversion: String

获取表示主机应用程序或 Exchange Server 的版本的字符串。

如果邮件外接程序在 Outlook 桌面客户端或 iOS 上运行, 则该`hostVersion`属性返回主机应用程序 (Outlook) 的版本。 在 Outlook 网页版中, 该属性返回的是 Exchange 服务器的版本。 一个示例是字符串 "15.0.468.0"。

##### <a name="type"></a>类型

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

#### <a name="owaview-string"></a>OWAView: String

获取表示 web 上的 Outlook 的当前视图的字符串。

返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。

如果主机应用程序不是 web 上的 Outlook, 则访问此属性将导致`undefined`。

Web 上的 Outlook 具有三个视图, 分别对应于屏幕的宽度和窗口, 以及可以显示的列数:

*   `OneColumn` 在屏幕较窄时显示。 Web 上的 Outlook 在智能手机的整个屏幕上使用此单列布局。
*   `TwoColumns` 在屏幕较宽时显示。 Outlook 网页版在大多数平板电脑上使用此视图。
*   `ThreeColumns` 在屏幕为宽屏时显示。 例如, web 上的 Outlook 在桌面计算机上的全屏窗口中使用此视图。

##### <a name="type"></a>类型

*   String

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|
