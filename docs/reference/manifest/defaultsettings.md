---
title: 清单文件中的 DefaultSettings 元素
description: 指定内容或任务窗格外接程序的默认源位置和其他默认设置。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: f86419bf6a3c18e3ae62091c53b1e8f82c706fb1
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149570"
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
