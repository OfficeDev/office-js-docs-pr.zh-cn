---
title: 清单文件中的授权元素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cc3b80e0e02eca9c197b82931a6f2011ba385d57
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477941"
---
# <a name="authorization-element"></a>Authorization 元素

指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。

**授权**是清单中[授权](authorizations.md)元素的子元素。

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
