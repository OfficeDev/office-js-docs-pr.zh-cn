---
title: 清单文件中的 Method 元素
description: Method 元素指定 Office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e3e74a73a3422a7789e82d6f0e7a516bd795ca8
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641323"
---
# <a name="method-element"></a>Method 元素

指定 office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。

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

## <a name="remarks"></a>说明

`Methods` `Method` 邮件外接程序不支持和元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

> [!IMPORTANT]
> 因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。 有关如何执行此操作的详细信息，请参阅[了解 Office JAVASCRIPT API](../../develop/understanding-the-javascript-api-for-office.md)。
