---
title: 清单文件中的 WebApplicationInfo 元素
description: Office 外接程序清单的 WebApplicationInfo 元素参考文档 (XML) 文件。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 037de49320a6d1a1ca7dce3446b4f4008a2f1331
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234161"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="bd0ae-103">WebApplicationInfo 元素</span><span class="sxs-lookup"><span data-stu-id="bd0ae-103">WebApplicationInfo element</span></span>

<span data-ttu-id="bd0ae-104">支持 Office 外接程序中的单一登录 (SSO)。此元素包含外接程序中的信息，如下所示：</span><span class="sxs-lookup"><span data-stu-id="bd0ae-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="bd0ae-105">Office 客户端应用程序可能需要权限的OAuth 2.0 资源。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-105">An OAuth 2.0 *resource* to which the Office client application might need permissions.</span></span>
- <span data-ttu-id="bd0ae-106">OAuth 2.0 *客户端*，可能需要访问 Microsoft Graph 的权限。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="bd0ae-107">Word、Excel、Outlook 和 PowerPoint 目前支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="bd0ae-108">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](../requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="bd0ae-109">如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="bd0ae-110">要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="bd0ae-111">**WebApplicationInfo** 是清单中的 [VersionOverrides](versionoverrides.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="bd0ae-112">子元素</span><span class="sxs-lookup"><span data-stu-id="bd0ae-112">Child elements</span></span>

|  <span data-ttu-id="bd0ae-113">元素</span><span class="sxs-lookup"><span data-stu-id="bd0ae-113">Element</span></span> |  <span data-ttu-id="bd0ae-114">必需</span><span class="sxs-lookup"><span data-stu-id="bd0ae-114">Required</span></span>  |  <span data-ttu-id="bd0ae-115">说明</span><span class="sxs-lookup"><span data-stu-id="bd0ae-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bd0ae-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="bd0ae-116">**Id**</span></span>    |  <span data-ttu-id="bd0ae-117">是</span><span class="sxs-lookup"><span data-stu-id="bd0ae-117">Yes</span></span>   |  <span data-ttu-id="bd0ae-118">在 Azure Active Directory v2.0 终结点中注册的加载项关联服务的 **应用程序 ID**。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="bd0ae-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="bd0ae-119">**MsaId**</span></span>    |  <span data-ttu-id="bd0ae-120">否</span><span class="sxs-lookup"><span data-stu-id="bd0ae-120">No</span></span>   |  <span data-ttu-id="bd0ae-121">加载项 Web 应用程序的 MSA 客户端 ID，如msm.live.com。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="bd0ae-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="bd0ae-122">**Resource**</span></span>  |  <span data-ttu-id="bd0ae-123">是</span><span class="sxs-lookup"><span data-stu-id="bd0ae-123">Yes</span></span>   |  <span data-ttu-id="bd0ae-124">指定在 Azure Active Directory v2.0 终结点中注册的加载项的 **应用程序 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="bd0ae-125">Scopes</span><span class="sxs-lookup"><span data-stu-id="bd0ae-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="bd0ae-126">是</span><span class="sxs-lookup"><span data-stu-id="bd0ae-126">Yes</span></span>  |  <span data-ttu-id="bd0ae-127">指定加载项对资源（如 Microsoft Graph）所需的权限。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="bd0ae-128">授权</span><span class="sxs-lookup"><span data-stu-id="bd0ae-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="bd0ae-129">否</span><span class="sxs-lookup"><span data-stu-id="bd0ae-129">No</span></span>   | <span data-ttu-id="bd0ae-130">指定加载项的 Web 应用程序需要授权的外部资源和所需的权限。</span><span class="sxs-lookup"><span data-stu-id="bd0ae-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="bd0ae-131">WebApplicationInfo 示例</span><span class="sxs-lookup"><span data-stu-id="bd0ae-131">WebApplicationInfo example</span></span>

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
