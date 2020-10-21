---
title: 清单文件中的 ExtendedPermissions 元素
description: 定义加载项访问关联的 Api 或功能所需的扩展权限的集合。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626440"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="3da83-103">ExtendedPermissions 元素</span><span class="sxs-lookup"><span data-stu-id="3da83-103">ExtendedPermissions element</span></span>

<span data-ttu-id="3da83-104">定义加载项访问关联的 Api 或功能所需的扩展权限的集合。</span><span class="sxs-lookup"><span data-stu-id="3da83-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="3da83-105">`ExtendedPermissions`元素是[VersionOverrides](versionoverrides.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="3da83-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3da83-106">对此元素的支持是在要求集1.9 中引入的。</span><span class="sxs-lookup"><span data-stu-id="3da83-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="3da83-107">请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="3da83-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="child-elements"></a><span data-ttu-id="3da83-108">子元素</span><span class="sxs-lookup"><span data-stu-id="3da83-108">Child elements</span></span>

|  <span data-ttu-id="3da83-109">元素</span><span class="sxs-lookup"><span data-stu-id="3da83-109">Element</span></span> |  <span data-ttu-id="3da83-110">必需</span><span class="sxs-lookup"><span data-stu-id="3da83-110">Required</span></span>  |  <span data-ttu-id="3da83-111">说明</span><span class="sxs-lookup"><span data-stu-id="3da83-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="3da83-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="3da83-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="3da83-113">否</span><span class="sxs-lookup"><span data-stu-id="3da83-113">No</span></span>   | <span data-ttu-id="3da83-114">定义外接程序访问关联的 API 或功能所需的扩展权限。</span><span class="sxs-lookup"><span data-stu-id="3da83-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="3da83-115">`ExtendedPermissions` 示例</span><span class="sxs-lookup"><span data-stu-id="3da83-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="3da83-116">以下是元素的示例 `ExtendedPermissions` 。</span><span class="sxs-lookup"><span data-stu-id="3da83-116">The following is an example of the `ExtendedPermissions` element.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a><span data-ttu-id="3da83-117">包含于</span><span class="sxs-lookup"><span data-stu-id="3da83-117">Contained in</span></span>

[<span data-ttu-id="3da83-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="3da83-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="3da83-119">可以包含</span><span class="sxs-lookup"><span data-stu-id="3da83-119">Can contain</span></span>

[<span data-ttu-id="3da83-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="3da83-120">ExtendedPermission</span></span>](extendedpermission.md)
