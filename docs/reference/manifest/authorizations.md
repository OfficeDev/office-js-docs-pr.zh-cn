---
title: 清单文件中的授权元素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 6a271423ddd549431c2f580e2793faab3c49090e
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477955"
---
# <a name="authorizations-element"></a><span data-ttu-id="98113-102">授权元素</span><span class="sxs-lookup"><span data-stu-id="98113-102">Authorizations element</span></span>

<span data-ttu-id="98113-103">指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。</span><span class="sxs-lookup"><span data-stu-id="98113-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="98113-104">**授权**是清单中的[WebApplicationInfo](webapplicationinfo.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="98113-104">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="98113-105">子元素</span><span class="sxs-lookup"><span data-stu-id="98113-105">Child elements</span></span>

|  <span data-ttu-id="98113-106">元素</span><span class="sxs-lookup"><span data-stu-id="98113-106">Element</span></span> |  <span data-ttu-id="98113-107">必需</span><span class="sxs-lookup"><span data-stu-id="98113-107">Required</span></span>  |  <span data-ttu-id="98113-108">说明</span><span class="sxs-lookup"><span data-stu-id="98113-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="98113-109">Authorization</span><span class="sxs-lookup"><span data-stu-id="98113-109">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="98113-110">是</span><span class="sxs-lookup"><span data-stu-id="98113-110">Yes</span></span>     |   <span data-ttu-id="98113-111">标识外接程序的 web 应用程序需要其授权的外部资源, 以及所需的范围 (权限)。</span><span class="sxs-lookup"><span data-stu-id="98113-111">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="98113-112">示例</span><span class="sxs-lookup"><span data-stu-id="98113-112">Example</span></span>

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
