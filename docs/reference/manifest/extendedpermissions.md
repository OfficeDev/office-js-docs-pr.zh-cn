---
title: 清单文件中的 ExtendedPermissions 元素
description: 定义加载项访问关联的 Api 或功能所需的扩展权限的集合。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: cf59d13d794f8f303da6cc0ca39066584bc3f56c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611531"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="bbcc4-103">ExtendedPermissions 元素</span><span class="sxs-lookup"><span data-stu-id="bbcc4-103">ExtendedPermissions element</span></span>

<span data-ttu-id="bbcc4-104">定义加载项访问关联的 Api 或功能所需的扩展权限的集合。</span><span class="sxs-lookup"><span data-stu-id="bbcc4-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="bbcc4-105">`ExtendedPermissions`元素是[VersionOverrides](versionoverrides.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="bbcc4-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bbcc4-106">此元素仅适用于针对 Exchange Online 的[Outlook 外接程序预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。</span><span class="sxs-lookup"><span data-stu-id="bbcc4-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="bbcc4-107">使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。</span><span class="sxs-lookup"><span data-stu-id="bbcc4-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bbcc4-108">子元素</span><span class="sxs-lookup"><span data-stu-id="bbcc4-108">Child elements</span></span>

|  <span data-ttu-id="bbcc4-109">元素</span><span class="sxs-lookup"><span data-stu-id="bbcc4-109">Element</span></span> |  <span data-ttu-id="bbcc4-110">必需</span><span class="sxs-lookup"><span data-stu-id="bbcc4-110">Required</span></span>  |  <span data-ttu-id="bbcc4-111">Description</span><span class="sxs-lookup"><span data-stu-id="bbcc4-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="bbcc4-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="bbcc4-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="bbcc4-113">否</span><span class="sxs-lookup"><span data-stu-id="bbcc4-113">No</span></span>   | <span data-ttu-id="bbcc4-114">定义外接程序访问关联的 API 或功能所需的扩展权限。</span><span class="sxs-lookup"><span data-stu-id="bbcc4-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="bbcc4-115">`ExtendedPermissions`示例</span><span class="sxs-lookup"><span data-stu-id="bbcc4-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="bbcc4-116">以下是元素的示例 `ExtendedPermissions` 。</span><span class="sxs-lookup"><span data-stu-id="bbcc4-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="bbcc4-117">包含于</span><span class="sxs-lookup"><span data-stu-id="bbcc4-117">Contained in</span></span>

[<span data-ttu-id="bbcc4-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="bbcc4-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="bbcc4-119">可以包含</span><span class="sxs-lookup"><span data-stu-id="bbcc4-119">Can contain</span></span>

[<span data-ttu-id="bbcc4-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="bbcc4-120">ExtendedPermission</span></span>](extendedpermission.md)
