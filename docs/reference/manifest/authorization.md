---
title: 清单文件中的授权元素
description: 指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: cece0934eb9db3175b173e97d7ab478827b7cda2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718438"
---
# <a name="authorization-element"></a><span data-ttu-id="5b38f-103">Authorization 元素</span><span class="sxs-lookup"><span data-stu-id="5b38f-103">Authorization element</span></span>

<span data-ttu-id="5b38f-104">指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。</span><span class="sxs-lookup"><span data-stu-id="5b38f-104">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>

<span data-ttu-id="5b38f-105">**授权**是清单中[授权](authorizations.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="5b38f-105">**Authorization** is a child element of the [Authorizations](authorizations.md) element in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5b38f-106">子元素</span><span class="sxs-lookup"><span data-stu-id="5b38f-106">Child elements</span></span>

|  <span data-ttu-id="5b38f-107">元素</span><span class="sxs-lookup"><span data-stu-id="5b38f-107">Element</span></span> |  <span data-ttu-id="5b38f-108">必需</span><span class="sxs-lookup"><span data-stu-id="5b38f-108">Required</span></span>  |  <span data-ttu-id="5b38f-109">说明</span><span class="sxs-lookup"><span data-stu-id="5b38f-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5b38f-110">**Resource**</span><span class="sxs-lookup"><span data-stu-id="5b38f-110">**Resource**</span></span>  |  <span data-ttu-id="5b38f-111">是</span><span class="sxs-lookup"><span data-stu-id="5b38f-111">Yes</span></span>   |  <span data-ttu-id="5b38f-112">指定外部资源的 URL。</span><span class="sxs-lookup"><span data-stu-id="5b38f-112">Specifies the URL of the external resource.</span></span>|
|  [<span data-ttu-id="5b38f-113">Scopes</span><span class="sxs-lookup"><span data-stu-id="5b38f-113">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="5b38f-114">是</span><span class="sxs-lookup"><span data-stu-id="5b38f-114">Yes</span></span>  |  <span data-ttu-id="5b38f-115">指定外接程序对资源所需的权限。</span><span class="sxs-lookup"><span data-stu-id="5b38f-115">Specifies the permissions that the add-in needs to the resource.</span></span>  |

## <a name="example"></a><span data-ttu-id="5b38f-116">示例</span><span class="sxs-lookup"><span data-stu-id="5b38f-116">Example</span></span>

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
