---
title: 清单文件中的授权元素
description: 指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 7ae0b9d0ec32a20846142a9fc89c48fe9cdf8053
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720657"
---
# <a name="authorizations-element"></a><span data-ttu-id="fa647-103">授权元素</span><span class="sxs-lookup"><span data-stu-id="fa647-103">Authorizations element</span></span>

<span data-ttu-id="fa647-104">指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。</span><span class="sxs-lookup"><span data-stu-id="fa647-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="fa647-105">**授权**是清单中的[WebApplicationInfo](webapplicationinfo.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="fa647-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="fa647-106">子元素</span><span class="sxs-lookup"><span data-stu-id="fa647-106">Child elements</span></span>

|  <span data-ttu-id="fa647-107">元素</span><span class="sxs-lookup"><span data-stu-id="fa647-107">Element</span></span> |  <span data-ttu-id="fa647-108">必需</span><span class="sxs-lookup"><span data-stu-id="fa647-108">Required</span></span>  |  <span data-ttu-id="fa647-109">说明</span><span class="sxs-lookup"><span data-stu-id="fa647-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fa647-110">Authorization</span><span class="sxs-lookup"><span data-stu-id="fa647-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="fa647-111">是</span><span class="sxs-lookup"><span data-stu-id="fa647-111">Yes</span></span>     |   <span data-ttu-id="fa647-112">标识外接程序的 web 应用程序需要其授权的外部资源，以及所需的范围（权限）。</span><span class="sxs-lookup"><span data-stu-id="fa647-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="fa647-113">示例</span><span class="sxs-lookup"><span data-stu-id="fa647-113">Example</span></span>

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
