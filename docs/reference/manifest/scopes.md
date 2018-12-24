---
title: 清单文件中的 Scopes 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 01d34481b14ac6a9186de07d352b9985dc1695a4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432639"
---
# <a name="scopes-element"></a><span data-ttu-id="c1f07-102">Scopes 元素</span><span class="sxs-lookup"><span data-stu-id="c1f07-102">Scopes element</span></span>

<span data-ttu-id="c1f07-103">包含加载项需要拥有的对 Microsoft Graph 的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c1f07-103">Contains permissions to Microsoft Graph that the add-in needs.</span></span> <span data-ttu-id="c1f07-104">Office 应用商店使用 Scopes 元素创建许可对话框。</span><span class="sxs-lookup"><span data-stu-id="c1f07-104">The Office Store uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="c1f07-105">当用户安装应用商店中的加载项时，系统会提示他们授予加载项对用户 Microsoft Graph 数据的指定访问权限。</span><span class="sxs-lookup"><span data-stu-id="c1f07-105">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c1f07-106">子元素</span><span class="sxs-lookup"><span data-stu-id="c1f07-106">Child elements</span></span>

|  <span data-ttu-id="c1f07-107">元素</span><span class="sxs-lookup"><span data-stu-id="c1f07-107">Element</span></span> |  <span data-ttu-id="c1f07-108">类型</span><span class="sxs-lookup"><span data-stu-id="c1f07-108">Type</span></span>  |  <span data-ttu-id="c1f07-109">说明</span><span class="sxs-lookup"><span data-stu-id="c1f07-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c1f07-110">**Scope**</span><span class="sxs-lookup"><span data-stu-id="c1f07-110">**Scope**</span></span>                |  <span data-ttu-id="c1f07-111">string</span><span class="sxs-lookup"><span data-stu-id="c1f07-111">string</span></span>     |   <span data-ttu-id="c1f07-112">Microsoft Graph 权限的名称，例如，Files.Read.All。</span><span class="sxs-lookup"><span data-stu-id="c1f07-112">The name of a permission to Microsoft Graph; for example, Files.Read.All.</span></span> |

## <a name="example"></a><span data-ttu-id="c1f07-113">示例</span><span class="sxs-lookup"><span data-stu-id="c1f07-113">Example</span></span>

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
