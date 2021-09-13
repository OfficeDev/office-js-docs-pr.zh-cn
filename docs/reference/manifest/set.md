---
title: 清单文件中的 Set 元素
description: Set 元素指定Office加载项Office所需的 JavaScript API 要求集。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 93524d64fd915d6f42f4e4a0cd0ab6cc3335f4ce
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152514"
---
# <a name="set-element"></a>Set 元素

指定加载项需要激活Office JavaScript API Office要求集。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a>包含于

[Sets](sets.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|Name|字符串|必需|[要求集](../../develop/office-versions-and-requirement-sets.md)名称。|
|MinVersion|字符串|可选|指定您的外接程序所需的 API 集的最低版本。 替代 **DefaultMinVersion** 的值（如果在父 Sets 元素 [中](sets.md) 指定）。|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

有关 **Set** 元素 **的 MinVersion** 属性和 **Sets** 元素 **的 DefaultMinVersion** 属性详细信息，请参阅在清单中设置 [Requirements 元素](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。

> [!IMPORTANT]
> 对于邮件外接程序，则只能使用一个 `"Mailbox"` 要求集。 此要求集包含 Outlook 邮件外接程序支持的整个 API 子集，你必须在邮件外接程序清单中指定 `"Mailbox"` 要求集（针对内容和任务窗格外接程序，非可选）。 另外，您无法在邮件外接程序中声明对特定方法的支持。
