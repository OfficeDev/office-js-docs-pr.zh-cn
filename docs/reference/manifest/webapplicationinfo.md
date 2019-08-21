---
title: 清单文件中的 WebApplicationInfo 元素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: e10aee1bf3fb99099d282acd428fa0348229701c
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477864"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="3e2df-102">WebApplicationInfo 元素</span><span class="sxs-lookup"><span data-stu-id="3e2df-102">WebApplicationInfo element</span></span>

<span data-ttu-id="3e2df-103">支持 Office 外接程序中的单一登录 (SSO)。此元素包含外接程序中的信息，如下所示：</span><span class="sxs-lookup"><span data-stu-id="3e2df-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="3e2df-104">OAuth 2.0 *资源*，Office 主机应用程序可能需要访问该资源的权限。</span><span class="sxs-lookup"><span data-stu-id="3e2df-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="3e2df-105">OAuth 2.0 *客户端*，可能需要访问 Microsoft Graph 的权限。</span><span class="sxs-lookup"><span data-stu-id="3e2df-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="3e2df-106">目前，Word、Excel、Outlook 和 PowerPoint 在预览版中支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="3e2df-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="3e2df-107">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="3e2df-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="3e2df-108">如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="3e2df-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="3e2df-109">要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。</span><span class="sxs-lookup"><span data-stu-id="3e2df-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="3e2df-110">**WebApplicationInfo** 是清单中的 [VersionOverrides](versionoverrides.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="3e2df-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="3e2df-111">子元素</span><span class="sxs-lookup"><span data-stu-id="3e2df-111">Child elements</span></span>

|  <span data-ttu-id="3e2df-112">元素</span><span class="sxs-lookup"><span data-stu-id="3e2df-112">Element</span></span> |  <span data-ttu-id="3e2df-113">必需</span><span class="sxs-lookup"><span data-stu-id="3e2df-113">Required</span></span>  |  <span data-ttu-id="3e2df-114">说明</span><span class="sxs-lookup"><span data-stu-id="3e2df-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3e2df-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="3e2df-115">**Id**</span></span>    |  <span data-ttu-id="3e2df-116">是</span><span class="sxs-lookup"><span data-stu-id="3e2df-116">Yes</span></span>   |  <span data-ttu-id="3e2df-117">在 Azure Active Directory v2.0 终结点中注册的加载项关联服务的**应用程序 ID**。</span><span class="sxs-lookup"><span data-stu-id="3e2df-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="3e2df-118">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="3e2df-118">**MsaId**</span></span>    |  <span data-ttu-id="3e2df-119">否</span><span class="sxs-lookup"><span data-stu-id="3e2df-119">No</span></span>   |  <span data-ttu-id="3e2df-120">在 msm.live.com 中注册的用于 MSA 的外接程序 web 应用程序的客户端 ID。</span><span class="sxs-lookup"><span data-stu-id="3e2df-120">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="3e2df-121">**Resource**</span><span class="sxs-lookup"><span data-stu-id="3e2df-121">**Resource**</span></span>  |  <span data-ttu-id="3e2df-122">是</span><span class="sxs-lookup"><span data-stu-id="3e2df-122">Yes</span></span>   |  <span data-ttu-id="3e2df-123">指定在 Azure Active Directory v2.0 终结点中注册的加载项的**应用程序 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="3e2df-123">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="3e2df-124">Scopes</span><span class="sxs-lookup"><span data-stu-id="3e2df-124">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="3e2df-125">是</span><span class="sxs-lookup"><span data-stu-id="3e2df-125">Yes</span></span>  |  <span data-ttu-id="3e2df-126">指定外接程序对资源所需的权限, 如 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="3e2df-126">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="3e2df-127">审核</span><span class="sxs-lookup"><span data-stu-id="3e2df-127">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="3e2df-128">否</span><span class="sxs-lookup"><span data-stu-id="3e2df-128">No</span></span>   | <span data-ttu-id="3e2df-129">指定加载项的 web 应用程序需要对其进行授权的外部资源以及所需的权限。</span><span class="sxs-lookup"><span data-stu-id="3e2df-129">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="3e2df-130">WebApplicationInfo 示例</span><span class="sxs-lookup"><span data-stu-id="3e2df-130">WebApplicationInfo example</span></span>

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
