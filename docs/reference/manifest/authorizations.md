---
title: 清单文件中的授权元素
description: 指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: 675585f99fc6261a2145219d553f02b9f9abded3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608752"
---
# <a name="authorizations-element"></a><span data-ttu-id="1f5b9-103">授权元素</span><span class="sxs-lookup"><span data-stu-id="1f5b9-103">Authorizations element</span></span>

<span data-ttu-id="1f5b9-104">指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。</span><span class="sxs-lookup"><span data-stu-id="1f5b9-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="1f5b9-105">**授权**是清单中的[WebApplicationInfo](webapplicationinfo.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="1f5b9-105">**Authorizations** is a child element of the [WebApplicationInfo](webapplicationinfo.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1f5b9-106">子元素</span><span class="sxs-lookup"><span data-stu-id="1f5b9-106">Child elements</span></span>

|  <span data-ttu-id="1f5b9-107">元素</span><span class="sxs-lookup"><span data-stu-id="1f5b9-107">Element</span></span> |  <span data-ttu-id="1f5b9-108">必需</span><span class="sxs-lookup"><span data-stu-id="1f5b9-108">Required</span></span>  |  <span data-ttu-id="1f5b9-109">Description</span><span class="sxs-lookup"><span data-stu-id="1f5b9-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1f5b9-110">Authorization</span><span class="sxs-lookup"><span data-stu-id="1f5b9-110">Authorization</span></span>](authorization.md)                |  <span data-ttu-id="1f5b9-111">是</span><span class="sxs-lookup"><span data-stu-id="1f5b9-111">Yes</span></span>     |   <span data-ttu-id="1f5b9-112">标识外接程序的 web 应用程序需要其授权的外部资源，以及所需的范围（权限）。</span><span class="sxs-lookup"><span data-stu-id="1f5b9-112">Identifies an external resource that the add-in's web application needs authorization to, and the scopes (permissions) that it needs.</span></span> |

## <a name="example"></a><span data-ttu-id="1f5b9-113">示例</span><span class="sxs-lookup"><span data-stu-id="1f5b9-113">Example</span></span>

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
