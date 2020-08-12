---
title: 清单文件中的 DefaultSettings 元素
description: 指定内容或任务窗格外接程序的默认源位置和其他默认设置。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a9711fb44390bcbda8979b8018eed1318c5579bc
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641464"
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

源位置和**DefaultSettings**元素中的其他设置仅适用于内容和任务窗格外接程序。对于邮件外接程序，您可以在[FormSettings](formsettings.md)元素中指定源文件和其他默认设置的默认位置。
