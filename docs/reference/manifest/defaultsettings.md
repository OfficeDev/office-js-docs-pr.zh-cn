---
title: 清单文件中的 DefaultSettings 元素
description: 指定内容或任务窗格外接程序的默认源位置和其他默认设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 11e398d86a702f4e45a5cea7b63e0380ce65d1749d0660789e96477744d73079
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095895"
---
# <a name="defaultsettings-element"></a>DefaultSettings 元素

指定内容或任务窗格外接程序的默认源位置和其他默认设置。

**外接程序类型：** 内容、任务窗格

## <a name="syntax"></a>语法

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

|元素|内容|邮件|任务窗格|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>注解

**DefaultSettings** 元素中的源位置和其他设置仅适用于内容和任务窗格外接程序。对于邮件外接程序，在 [FormSettings](formsettings.md)元素中指定源文件的默认位置和其他默认设置。
