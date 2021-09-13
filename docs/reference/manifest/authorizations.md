---
title: 清单文件中 Authorizations 元素
description: 指定加载项的 Web 应用程序需要授权的外部资源和所需权限。
ms.date: 08/12/2019
ms.localizationpriority: medium
ms.openlocfilehash: 4b13e26f13fae6fefd579868df8b67dd94cb35c4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151927"
---
# <a name="authorizations-element"></a>Authorizations 元素

指定加载项的 Web 应用程序需要授权的外部资源和所需权限。

**授权** 是清单中 [WebApplicationInfo](webapplicationinfo.md) 元素的子元素。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Authorization](authorization.md)                |  是     |   标识外接程序的 Web 应用程序需要授权到的外部资源，以及范围 (所需的) 权限。 |

## <a name="example"></a>示例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
