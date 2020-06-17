---
title: 清单文件中的 Scopes 元素
description: Scope 元素包含加载项连接到外部资源所需的权限。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: be68033e86de736703d9d1593ad361918d5a147d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612239"
---
# <a name="scopes-element"></a><span data-ttu-id="dbd36-103">Scopes 元素</span><span class="sxs-lookup"><span data-stu-id="dbd36-103">Scopes element</span></span>

<span data-ttu-id="dbd36-104">包含外接程序需要外部资源的权限，如 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="dbd36-104">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="dbd36-105">当 Microsoft Graph 是资源时，AppSource 使用 Scope 元素创建同意对话框。</span><span class="sxs-lookup"><span data-stu-id="dbd36-105">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="dbd36-106">当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。</span><span class="sxs-lookup"><span data-stu-id="dbd36-106">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="dbd36-107">**作用域**是清单中的[WebApplicationInfo](webapplicationinfo.md)和[授权](authorization.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="dbd36-107">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dbd36-108">子元素</span><span class="sxs-lookup"><span data-stu-id="dbd36-108">Child elements</span></span>

|  <span data-ttu-id="dbd36-109">元素</span><span class="sxs-lookup"><span data-stu-id="dbd36-109">Element</span></span> |  <span data-ttu-id="dbd36-110">必需</span><span class="sxs-lookup"><span data-stu-id="dbd36-110">Required</span></span>  |  <span data-ttu-id="dbd36-111">Description</span><span class="sxs-lookup"><span data-stu-id="dbd36-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="dbd36-112">**Scope**</span><span class="sxs-lookup"><span data-stu-id="dbd36-112">**Scope**</span></span>                |  <span data-ttu-id="dbd36-113">是</span><span class="sxs-lookup"><span data-stu-id="dbd36-113">Yes</span></span>     |   <span data-ttu-id="dbd36-114">权限的名称;例如，Files. All 或 profile。</span><span class="sxs-lookup"><span data-stu-id="dbd36-114">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="dbd36-115">示例</span><span class="sxs-lookup"><span data-stu-id="dbd36-115">Example</span></span>

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
