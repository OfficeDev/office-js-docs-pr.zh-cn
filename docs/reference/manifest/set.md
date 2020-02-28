---
title: 清单文件中的 Set 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324804"
---
# <a name="set-element"></a>Set 元素

指定 office 外接程序需要激活的 Office JavaScript API 中的要求集。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>包含于

[Sets](sets.md)

## <a name="attributes"></a>属性

|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|名称|string|必需|[要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)名称。|
|MinVersion|字符串|可选|指定您的外接程序所需的 API 集的最低版本。 重写**DefaultMinVersion**的值（如果它在父[集](sets.md)元素中指定）。|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

有关**Set**元素的**MinVersion**属性和**Sets**元素的**DefaultMinVersion**属性的详细信息，请参阅[在清单中设置需求元素](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。

> [!IMPORTANT] 
> 对于邮件外接程序，则只能使用一个 `"Mailbox"` 要求集。 此要求集包含 Outlook 邮件外接程序支持的整个 API 子集，你必须在邮件外接程序清单中指定 `"Mailbox"` 要求集（针对内容和任务窗格外接程序，非可选）。 另外，您无法在邮件外接程序中声明对特定方法的支持。
