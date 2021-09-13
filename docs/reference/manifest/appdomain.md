---
title: 清单文件中的 AppDomain 元素
description: 指定您的外接程序使用且应受外接程序信任的其他Office。
ms.date: 06/12/2020
ms.localizationpriority: medium
ms.openlocfilehash: c17195e6d9d3f4f22465c8aa1fc626afd3eb06c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151931"
---
# <a name="appdomain-element"></a>AppDomain 元素

指定除[SourceLocation](sourcelocation.md)Office中指定的域之外，还应信任其他域。 指定域具有以下效果：

- 它使页面中、路由或域中的其他资源可以直接在桌面和平台加载项的根任务窗格中Office打开。  (在 **AppDomain** 中指定域对于 Office web 版 或在 IFrame 中打开资源不是必需的，也不需要在对话框 [API](../../develop/dialog-api-in-office-add-ins.md).) 打开的对话框中打开资源
- 它使域中的页面能够Office.js内 IFrame 进行 API 调用。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain.com</AppDomain>`）。
> 2. 如果存在域的显式端口，请 (例如 `<AppDomain>https://myappdomain.com:9999</AppDomain>` ，) 。
> 3. 如果需要信任子域，请 (，例如 `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`) 。 子域和 `mysubdomain.mydomain.com` `mydomain.com` 是不同的域。 如果两者都需要受信任，则两者都需要位于单独的 **AppDomain** 元素中。
> 4. 列出与 [SourceLocation](sourcelocation.md) 元素中指定的域相同的域没有任何效果，并且可能会令人误解。 特别是，在 进行开发时， `localhost` 不需要为 **创建 AppDomain** 元素 `localhost` 。
> 5. 请勿包含通过域的 URL 的任何段。 例如，不包括页面的完整 URL。
> 6. 不要 *将* 结束斜杠"/"放在值上。

## <a name="contained-in"></a>包含于

[AppDomains](appdomains.md)

## <a name="remarks"></a>注解

有关详细信息，请参阅 [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)。
