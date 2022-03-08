---
title: 清单文件中的 Scopes 元素
description: Scopes 元素包含外接程序连接到外部资源所需的权限。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 883a1e318df7262bf8cdbd9d97b9d02d201066d8
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340398"
---
# <a name="scopes-element"></a>Scopes 元素

包含外接程序对外部资源（如 Microsoft Graph）所需的权限。 当 Microsoft Graph 资源时，AppSource 使用 Scopes 元素创建同意对话框。 当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。

**外接程序类型：** 任务窗格、邮件、内容

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 内容 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)

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
