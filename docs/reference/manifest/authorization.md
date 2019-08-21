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
# <a name="authorization-element"></a><span data-ttu-id="f52ec-102">Authorization 元素</span><span class="sxs-lookup"><span data-stu-id="f52ec-102">Authorization element</span></span>

<span data-ttu-id="f52ec-103">指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。</span><span class="sxs-lookup"><span data-stu-id="f52ec-103">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="f52ec-104">**授权**是清单中[授权](authorizations.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="f52ec-104">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f52ec-105">子元素</span><span class="sxs-lookup"><span data-stu-id="f52ec-105">Child elements</span></span>

|  <span data-ttu-id="f52ec-106">元素</span><span class="sxs-lookup"><span data-stu-id="f52ec-106">Element</span></span> |  <span data-ttu-id="f52ec-107">必需</span><span class="sxs-lookup"><span data-stu-id="f52ec-107">Required</span></span>  |  <span data-ttu-id="f52ec-108">说明</span><span class="sxs-lookup"><span data-stu-id="f52ec-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f52ec-109">**Resource**</span><span class="sxs-lookup"><span data-stu-id="f52ec-109">**Resource**</span></span>  |  <span data-ttu-id="f52ec-110">是</span><span class="sxs-lookup"><span data-stu-id="f52ec-110">Yes</span></span>   |  <span data-ttu-id="f52ec-111">指定外部资源的 URL。</span><span class="sxs-lookup"><span data-stu-id="f52ec-111">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="f52ec-112">Scopes</span><span class="sxs-lookup"><span data-stu-id="f52ec-112">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="f52ec-113">是</span><span class="sxs-lookup"><span data-stu-id="f52ec-113">Yes</span></span>  |  <span data-ttu-id="f52ec-114">指定外接程序对资源所需的权限。</span><span class="sxs-lookup"><span data-stu-id="f52ec-114">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="f52ec-115">示例</span><span class="sxs-lookup"><span data-stu-id="f52ec-115">Example</span></span>

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
