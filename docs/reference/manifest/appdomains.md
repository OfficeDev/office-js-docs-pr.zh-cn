---
title: 清单文件中的 AppDomains 元素
description: 列出除外接程序将使用的 元素中指定的域Office且应受用户信任 `SourceLocation` 的域Office。
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: 6bf1785cf11e31648d9bc69e101cd5a5cf3ecb9f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151928"
---
# <a name="appdomains-element"></a>AppDomains 元素

列出除 元素中指定的域之外，你的外接程序将使用Office且应受用户信任的任何域 `SourceLocation` Office。 这使域中的页面能够调用Office.js内 IFrame 的 API，并产生其他效果。 对于每个其他域，指定 **AppDomain** 元素。

 **外接程序类型：** 内容、任务窗格、邮件

## <a name="syntax"></a>语法

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> **AppDomain** 元素的值存在一些限制。 有关详细信息，请参阅 [AppDomain](appdomain.md)。

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>可以包含

[AppDomain](appdomain.md)

## <a name="remarks"></a>注释

默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。 此元素不能为空。
