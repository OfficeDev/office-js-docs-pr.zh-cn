---
title: 清单文件中的 Set 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0f137f7b08d6f1d0b0d972173c8085713b0f979d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432765"
---
# <a name="set-element"></a>Set 元素

指定来自适用于 Office 的 JavaScript API 的要求集，Office 外接程序需要该集才能激活。

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
|名称|字符串|必需|[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)名称。|
|MinVersion|字符串|可选|指定您的外接程序所需的 API 集的最低版本。如果 **DefaultMinVersion** 的值已在父 [Sets](sets.md) 元素中指定，则替代该值。|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置 Requirements 元素](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)。

> [!IMPORTANT] 
> 对于邮件外接程序，则只能使用一个 `"Mailbox"` 要求集。 此要求集包含 Outlook 邮件外接程序支持的整个 API 子集，你必须在邮件外接程序清单中指定 `"Mailbox"` 要求集（针对内容和任务窗格外接程序，非可选）。 另外，无法在邮件外接程序中声明对特定方法的支持。
