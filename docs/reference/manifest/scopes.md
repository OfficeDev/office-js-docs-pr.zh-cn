---
title: 清单文件中的 Scopes 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdc9ebeb6fe4167a5ed5e9407f6ecc82d5b8d507
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771784"
---
# <a name="scopes-element"></a><span data-ttu-id="6519a-102">Scopes 元素</span><span class="sxs-lookup"><span data-stu-id="6519a-102">Scopes element</span></span>

<span data-ttu-id="6519a-103">包含加载项需要拥有的对 Microsoft Graph 的访问权限。</span><span class="sxs-lookup"><span data-stu-id="6519a-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="6519a-104">AppSource 使用 Scope 元素创建同意对话框。</span><span class="sxs-lookup"><span data-stu-id="6519a-104">AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="6519a-105">当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。</span><span class="sxs-lookup"><span data-stu-id="6519a-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6519a-106">子元素</span><span class="sxs-lookup"><span data-stu-id="6519a-106">Child elements</span></span>

|  <span data-ttu-id="6519a-107">元素</span><span class="sxs-lookup"><span data-stu-id="6519a-107">Element</span></span> |  <span data-ttu-id="6519a-108">类型</span><span class="sxs-lookup"><span data-stu-id="6519a-108">Type</span></span>  |  <span data-ttu-id="6519a-109">说明</span><span class="sxs-lookup"><span data-stu-id="6519a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6519a-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="6519a-110">**Scope**</span></span>                |  <span data-ttu-id="6519a-111">string</span><span class="sxs-lookup"><span data-stu-id="6519a-111">string</span></span>     |   <span data-ttu-id="6519a-112">Microsoft Graph 权限的名称，例如，Files.Read.All。</span><span class="sxs-lookup"><span data-stu-id="6519a-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="6519a-113">示例</span><span class="sxs-lookup"><span data-stu-id="6519a-113">Example</span></span>

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
