---
title: 清单文件中的 WebApplicationInfo 元素
description: Office 外接程序清单的 WebApplicationInfo 元素的参考文档 (XML) 文件。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 8644529d82204cb9fbc07c6fe9f8a35b60a512c8
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293805"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="1afd1-103">WebApplicationInfo 元素</span><span class="sxs-lookup"><span data-stu-id="1afd1-103">WebApplicationInfo element</span></span>

<span data-ttu-id="1afd1-104">支持 Office 外接程序中的单一登录 (SSO)。此元素包含外接程序中的信息，如下所示：</span><span class="sxs-lookup"><span data-stu-id="1afd1-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="1afd1-105">Office 客户端应用程序可能需要其权限的 OAuth 2.0 *资源* 。</span><span class="sxs-lookup"><span data-stu-id="1afd1-105">An OAuth 2.0 *resource* to which the Office client application might need permissions.</span></span>
- <span data-ttu-id="1afd1-106">OAuth 2.0 *客户端*，可能需要访问 Microsoft Graph 的权限。</span><span class="sxs-lookup"><span data-stu-id="1afd1-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="1afd1-107">目前，Word、Excel、Outlook 和 PowerPoint 支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="1afd1-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="1afd1-108">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="1afd1-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="1afd1-109">如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="1afd1-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="1afd1-110">要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。</span><span class="sxs-lookup"><span data-stu-id="1afd1-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="1afd1-111">**WebApplicationInfo** 是清单中的 [VersionOverrides](versionoverrides.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="1afd1-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="1afd1-112">子元素</span><span class="sxs-lookup"><span data-stu-id="1afd1-112">Child elements</span></span>

|  <span data-ttu-id="1afd1-113">元素</span><span class="sxs-lookup"><span data-stu-id="1afd1-113">Element</span></span> |  <span data-ttu-id="1afd1-114">必需</span><span class="sxs-lookup"><span data-stu-id="1afd1-114">Required</span></span>  |  <span data-ttu-id="1afd1-115">说明</span><span class="sxs-lookup"><span data-stu-id="1afd1-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1afd1-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="1afd1-116">**Id**</span></span>    |  <span data-ttu-id="1afd1-117">是</span><span class="sxs-lookup"><span data-stu-id="1afd1-117">Yes</span></span>   |  <span data-ttu-id="1afd1-118">在 Azure Active Directory v2.0 终结点中注册的加载项关联服务的**应用程序 ID**。</span><span class="sxs-lookup"><span data-stu-id="1afd1-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="1afd1-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="1afd1-119">**MsaId**</span></span>    |  <span data-ttu-id="1afd1-120">否</span><span class="sxs-lookup"><span data-stu-id="1afd1-120">No</span></span>   |  <span data-ttu-id="1afd1-121">在 msm.live.com 中注册的用于 MSA 的外接程序 web 应用程序的客户端 ID。</span><span class="sxs-lookup"><span data-stu-id="1afd1-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="1afd1-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="1afd1-122">**Resource**</span></span>  |  <span data-ttu-id="1afd1-123">是</span><span class="sxs-lookup"><span data-stu-id="1afd1-123">Yes</span></span>   |  <span data-ttu-id="1afd1-124">指定在 Azure Active Directory v2.0 终结点中注册的加载项的**应用程序 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="1afd1-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="1afd1-125">Scopes</span><span class="sxs-lookup"><span data-stu-id="1afd1-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="1afd1-126">是</span><span class="sxs-lookup"><span data-stu-id="1afd1-126">Yes</span></span>  |  <span data-ttu-id="1afd1-127">指定外接程序对资源所需的权限，如 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="1afd1-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="1afd1-128">授权</span><span class="sxs-lookup"><span data-stu-id="1afd1-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="1afd1-129">否</span><span class="sxs-lookup"><span data-stu-id="1afd1-129">No</span></span>   | <span data-ttu-id="1afd1-130">指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。</span><span class="sxs-lookup"><span data-stu-id="1afd1-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="1afd1-131">WebApplicationInfo 示例</span><span class="sxs-lookup"><span data-stu-id="1afd1-131">WebApplicationInfo example</span></span>

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
