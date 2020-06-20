---
title: 清单文件中的 AppDomain 元素
description: 指定加载项使用的其他域，并且应受 Office 信任。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778646"
---
# <a name="appdomain-element"></a>AppDomain 元素

指定除了在[SourceLocation 元素](sourcelocation.md)中指定的域之外，Office 应信任的其他域。 指定域具有以下效果：

- 它允许在桌面 Office 平台上的加载项的根任务窗格中直接打开页面、路由或域中的其他资源。 （为 web 上的 Office**指定不需要的域**或在 IFrame 中打开资源，也需要在使用[对话框 API](../../develop/dialog-api-in-office-add-ins.md)打开的对话框中打开资源时。）
- 它使域中的页面可以从加载项中的 Iframe 进行 Office.js API 调用。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. **AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain.com</AppDomain>`）。
> 2. 如果有域的显式端口，请将其包括在内（例如， `<AppDomain>https://myappdomain.com:9999</AppDomain>` ）。
> 3. 如果需要信任某个子域，请将其包括在内（例如， `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ）。 子域 `mysubdomain.mydomain.com` 和 `mydomain.com` 不同的域。 如果两者都需要信任，则这两个元素都需要位于单独的**AppDomain**元素中。
> 4. 列出与[SourceLocation 元素](sourcelocation.md)中指定的域相同的域不起作用，并且可能会引起误导。 特别是在上进行开发时 `localhost` ，不需要为创建**AppDomain**元素 `localhost` 。
> 5. 不要将任何段的 URL 包含在域之后。 例如，不要包含页面的完整 URL。
> 6. 不要*在值上添加一个*结束斜杠 "/"。

## <a name="contained-in"></a>包含于

[AppDomains](appdomains.md)

## <a name="remarks"></a>注解

有关详细信息，请参阅 [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)。
