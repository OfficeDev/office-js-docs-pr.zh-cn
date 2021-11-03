---
title: 清单文件中的 Scopes 元素
description: Scopes 元素包含外接程序连接到外部资源所需的权限。
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: 16e8a19a7aa73efa6aac00c915fde8d2b8647bad
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681533"
---
# <a name="scopes-element"></a>Scopes 元素

包含外接程序对外部资源（如 Microsoft Graph）所需的权限。 当 Microsoft Graph 是资源时，AppSource 使用 Scopes 元素创建同意对话框。 当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。

**Scopes** 是清单 [中 WebApplicationInfo](webapplicationinfo.md) 元素的子元素。

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
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc<Resource>
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
