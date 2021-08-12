---
title: 清单文件中的授权元素
description: 指定加载项的 Web 应用程序需要授权的外部资源以及所需的权限。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: af40a47c4ae30b6d18d3457704487027ff18ac92da2a3ae23cf1afe5c1e9b46a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087707"
---
# <a name="authorization-element"></a>Authorization 元素

指定加载项的 Web 应用程序需要授权的外部资源和所需权限。

**授权** 是清单中 [Authorizations](authorizations.md) 元素的子元素。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Resource**  |  是   |  指定外部资源的 URL。|
|  [Scopes](scopes.md)                |  是  |  指定外接程序对资源所需的权限。  |

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
