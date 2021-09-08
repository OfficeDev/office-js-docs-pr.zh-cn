---
title: 清单文件中的 Sets 元素
description: Sets 元素指定外接程序Office激活Office JavaScript API 的最小集合。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938660"
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

