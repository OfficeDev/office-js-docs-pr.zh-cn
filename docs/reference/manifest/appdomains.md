---
title: 清单文件中的 AppDomains 元素
description: 列出除 Office 外接程序将使用的元素中指定的域之外的所有域 `SourceLocation` ，以及 office 应信任的域。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778653"
---
# <a name="appdomains-element"></a>AppDomains 元素

列出 `SourceLocation` 您的 Office 外接程序将使用且应受 office 信任的任何域（除了元素中指定的域）。 这使域中的页面可以调用来自加载项中的 Iframe 的 Office.js Api，并具有其他效果。 对于每个其他域，指定 **AppDomain** 元素。

 **外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> 对可以成为**AppDomain**元素的值的值有一些限制。 有关详细信息，请参阅[AppDomain](appdomain.md)。

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

[AppDomain](appdomain.md)

## <a name="remarks"></a>注释

默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。 此元素不能为空。
