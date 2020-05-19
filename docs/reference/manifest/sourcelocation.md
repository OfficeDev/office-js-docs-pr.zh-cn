---
title: 清单文件中的 SourceLocation 元素
description: SourceLocation 元素指定 Office 外接程序的源文件位置。
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 642780c3231523ea579ca548b3f3f984b2856666
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278397"
---
# <a name="sourcelocation-element"></a>SourceLocation 元素

将 Office 外接程序的源文件位置指定为一个长度介于1到2018个字符之间的 URL。 源位置必须是 HTTPS 地址，而非文件路径。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a>包含于

- [DefaultSettings](defaultsettings.md)（内容和任务窗格外接程序）
- [FormSettings](formsettings.md)（邮件外接程序）
- [ExtensionPoint](extensionpoint.md) （上下文和 LaunchEvent （预览）邮件外接程序）

## <a name="can-contain"></a>可以包含

[Override](override.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**描述**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|指定该设置的默认值，表示为 [DefaultLocale](defaultlocale.md) 元素中指定的区域设置。|
