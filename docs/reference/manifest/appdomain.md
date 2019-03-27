---
title: 清单文件中的 AppDomain 元素
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: 8216603c87a7dcafde84d25a82f068c9aa86ed96
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870406"
---
# <a name="appdomain-element"></a>AppDomain 元素

指定将用于在外接程序窗口中加载页面的其他域。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain</AppDomain>`）。
> 2. 不要** 在值上添加一个结束斜杠 "/"。

## <a name="contained-in"></a>包含于

[AppDomains](appdomains.md)

## <a name="remarks"></a>注释

**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。 有关详细信息，请参阅 [Office 加载项 XML 清单](/office/dev/add-ins/develop/add-in-manifests)。
