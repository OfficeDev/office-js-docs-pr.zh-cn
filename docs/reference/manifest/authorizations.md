---
title: 清单文件中 Authorizations 元素
description: 指定加载项的 Web 应用程序需要授权的外部资源和所需权限。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 068e6753e2e8e947e5e6e3c0885e7cd006165660862a37346eea114abb81a9b8
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092497"
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
