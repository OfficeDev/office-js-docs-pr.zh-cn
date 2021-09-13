---
title: 清单文件中的 Sets 元素
description: Sets 元素指定外接程序Office激活Office JavaScript API 的最小集合。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 38707ec78a79e9104dd21f9fa5ceab8c6fbd2c79
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152508"
---
# <a name="sets-element"></a>Sets 元素

指定您的外接程序Office激活所需的 Office JavaScript API 的最小子集。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>包含于

[要求](requirements.md)

## <a name="can-contain"></a>可以包含

[Set](set.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|字符串|可选|指定所有子 **Set** 元素的默认 MinVersion [属性值。](set.md) 默认值为“1.1”。|

## <a name="remarks"></a>注释

有关要求集的详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

有关 **Set** 元素 **的 MinVersion** 属性和 **Sets** 元素 **的 DefaultMinVersion** 属性详细信息，请参阅在清单中设置 [Requirements 元素](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。

