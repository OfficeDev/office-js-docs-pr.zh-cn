---
title: 清单文件中的 Scopes 元素
description: Scopes 元素包含外接程序连接到外部资源所需的权限。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 05582ae05c13fae8e2272de3fe6111c5ff639f938a817fd0b50ad22e4234d033
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087252"
---
# <a name="scopes-element"></a>Scopes 元素

包含外接程序对外部资源（如 Microsoft Graph）所需的权限。 当 Microsoft Graph 是资源时，AppSource 使用 Scopes 元素创建同意对话框。 当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。

**Scopes** 是清单中 [WebApplicationInfo](webapplicationinfo.md) 和 [Authorization](authorization.md) 元素的子元素。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Scope**                |  是     |   权限的名称;例如，Files.Read.All 或 profile。 |

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
