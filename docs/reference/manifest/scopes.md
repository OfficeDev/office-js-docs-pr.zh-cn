---
title: 清单文件中的 Scopes 元素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 1e36bdcd0cdcaa8c842e924c2543d56bdc4e26a7
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477731"
---
# <a name="scopes-element"></a>Scopes 元素

包含外接程序需要外部资源的权限, 如 Microsoft Graph。 当 Microsoft Graph 是资源时, AppSource 使用 Scope 元素创建同意对话框。 当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。

**作用域**是清单中的[WebApplicationInfo](webapplicationinfo.md)和[授权](authorization.md)元素的子元素。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Scope**                |  是     |   权限的名称;例如, Files. All 或 profile。 |

## <a name="example"></a>示例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
