---
title: 清单文件中的 SourceLocation 元素
description: SourceLocation 元素指定外接程序的Office位置。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4dcd093db2f23220eaa34c0c81300c4994c1a697
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590895"
---
# <a name="sourcelocation-element"></a>SourceLocation 元素

指定外接程序的源文件位置Office 1 到 2018 个字符之间的 URL。 源位置必须是 HTTPS 地址，而非文件路径。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>包含于

- [DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）
- [FormSettings](formsettings.md)（邮件外接程序）
- [ExtensionPoint](extensionpoint.md) (上下文和 LaunchEvent 邮件外接程序) 

## <a name="can-contain"></a>可以包含

[Override](override.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|
