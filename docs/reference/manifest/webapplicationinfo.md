---
title: 清单文件中的 WebApplicationInfo 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ae1f2ea43e795d8f4ac2f634785fc69b0ca7a3d0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457626"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="e4d7a-102">WebApplicationInfo 元素</span><span class="sxs-lookup"><span data-stu-id="e4d7a-102">WebApplicationInfo element</span></span>

<span data-ttu-id="e4d7a-103">支持 Office 外接程序中的单一登录 (SSO)。此元素包含外接程序中的信息，如下所示：</span><span class="sxs-lookup"><span data-stu-id="e4d7a-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="e4d7a-104">OAuth 2.0 *资源*，Office 主机应用程序可能需要访问该资源的权限。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="e4d7a-105">OAuth 2.0 *客户端*，可能需要访问 Microsoft Graph 的权限。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="e4d7a-106">目前，Word、Excel、Outlook 和 PowerPoint 在预览版中支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="e4d7a-107">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="e4d7a-108">如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="e4d7a-109">要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="e4d7a-110">**WebApplicationInfo** 是清单中的 [VersionOverrides](versionoverrides.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="e4d7a-111">子元素</span><span class="sxs-lookup"><span data-stu-id="e4d7a-111">Child elements</span></span>

|  <span data-ttu-id="e4d7a-112">元素</span><span class="sxs-lookup"><span data-stu-id="e4d7a-112">Element</span></span> |  <span data-ttu-id="e4d7a-113">必需</span><span class="sxs-lookup"><span data-stu-id="e4d7a-113">Required</span></span>  |  <span data-ttu-id="e4d7a-114">说明</span><span class="sxs-lookup"><span data-stu-id="e4d7a-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e4d7a-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="e4d7a-115">**Id**</span></span>    |  <span data-ttu-id="e4d7a-116">是</span><span class="sxs-lookup"><span data-stu-id="e4d7a-116">Yes</span></span>   |  <span data-ttu-id="e4d7a-117">在 Azure Active Directory v2.0 终结点中注册的加载项关联服务的**应用程序 ID**。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="e4d7a-118">**Resource**</span><span class="sxs-lookup"><span data-stu-id="e4d7a-118">**Resource**</span></span>  |  <span data-ttu-id="e4d7a-119">是</span><span class="sxs-lookup"><span data-stu-id="e4d7a-119">Yes</span></span>   |  <span data-ttu-id="e4d7a-120">指定在 Azure Active Directory v2.0 终结点中注册的加载项的**应用程序 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-120">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="e4d7a-121">Scopes</span><span class="sxs-lookup"><span data-stu-id="e4d7a-121">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="e4d7a-122">否</span><span class="sxs-lookup"><span data-stu-id="e4d7a-122">No</span></span>  |  <span data-ttu-id="e4d7a-123">指定加载项需要拥有的对 Microsoft Graph 的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-123">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="e4d7a-124">目前，加载项的 Resource 必须与其 Host 一致。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-124">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="e4d7a-125">Office 不会请求获取加载项令牌，除非可以证明所有权。目前，这是通过在 Resource 的完全限定的域名下托管加载项来完成。</span><span class="sxs-lookup"><span data-stu-id="e4d7a-125">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="e4d7a-126">WebApplicationInfo 示例</span><span class="sxs-lookup"><span data-stu-id="e4d7a-126">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>        
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
