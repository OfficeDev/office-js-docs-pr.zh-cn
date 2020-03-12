---
title: 清单文件中的 AppDomain 元素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: da28b3b4dec5d669462a781db3c0628bd32c7182
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596786"
---
# <a name="appdomain-element"></a>AppDomain 元素

指定在外接程序窗口中加载页面的其他域。 此外，它还列出了可以从加载项内的 Iframe 中进行的 Office .js API 调用的受信任域。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain</AppDomain>`）。
> 2. 不要*在值上添加一个*结束斜杠 "/"。

## <a name="contained-in"></a>包含于

[AppDomains](appdomains.md)

## <a name="remarks"></a>注释

**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。 有关详细信息，请参阅 [Office 加载项 XML 清单](../../develop/add-in-manifests.md)。
