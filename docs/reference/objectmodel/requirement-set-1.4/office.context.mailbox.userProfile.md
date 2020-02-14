---
title: Office.context.mailbox.userProfile - 要求集 1.4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0532a9971a05412d37334f4c5a4b6b12654f61f3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950990"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile

提供有关 Outlook 外接程序中的用户的信息。

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

## <a name="properties"></a>属性

| 属性 | 最低<br>权限级别 | 型号 | 返回类型 | 最低<br>要求集 |
|---|---|---|---|:---:|
| [displayName](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | ReadItem | 撰写<br>读取 | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [emailAddress](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | ReadItem | 撰写<br>读取 | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [时区](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | ReadItem | 撰写<br>读取 | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
