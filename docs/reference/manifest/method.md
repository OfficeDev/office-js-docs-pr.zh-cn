---
title: 清单文件中的 Method 元素
description: Method 元素指定加载项激活Office JavaScript API Office所需的单个方法。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 811cd84e1ad2aade8b7042eefa822eee6b2ab200a8fa1b71c9fe5fc34874ec66
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089724"
---
# <a name="method-element"></a>Method 元素

指定加载项激活Office JavaScript API Office JavaScript API 中的单个方法。

**外接程序类型：** 内容、任务窗格

## <a name="syntax"></a>语法

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>包含于

[Methods](methods.md)

## <a name="attributes"></a>属性

|属性|类型|必需|说明|
|:-----|:-----|:-----|:-----|
|Name|字符串|必需|指定由其父对象限定的所需方法的名称。 例如，若要指定 `getSelectedDataAsync` 方法，必须指定 `"Document.getSelectedDataAsync"` 。|

## <a name="remarks"></a>备注

`Methods`邮件外接程序不支持 和 `Method` 元素。有关要求集详细信息，请参阅Office[版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

> [!IMPORTANT]
> 因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。 若要详细了解如何操作，请参阅了解 Office [JavaScript API。](../../develop/understanding-the-javascript-api-for-office.md)
