---
title: 清单文件中的 ExtendedPermissions 元素
description: 定义加载项访问关联的 Api 或功能所需的扩展权限的集合。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 86d898052af6ba0e6f6bc8b341fff9f0f8408967
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718221"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="36d55-103">ExtendedPermissions 元素</span><span class="sxs-lookup"><span data-stu-id="36d55-103">ExtendedPermissions element</span></span>

<span data-ttu-id="36d55-104">定义加载项访问关联的 Api 或功能所需的扩展权限的集合。</span><span class="sxs-lookup"><span data-stu-id="36d55-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="36d55-105">`ExtendedPermissions`元素是[VersionOverrides](versionoverrides.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="36d55-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="36d55-106">此元素仅适用于针对 Exchange Online 的[Outlook 外接程序预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。</span><span class="sxs-lookup"><span data-stu-id="36d55-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="36d55-107">使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。</span><span class="sxs-lookup"><span data-stu-id="36d55-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="36d55-108">子元素</span><span class="sxs-lookup"><span data-stu-id="36d55-108">Child elements</span></span>

|  <span data-ttu-id="36d55-109">元素</span><span class="sxs-lookup"><span data-stu-id="36d55-109">Element</span></span> |  <span data-ttu-id="36d55-110">必需</span><span class="sxs-lookup"><span data-stu-id="36d55-110">Required</span></span>  |  <span data-ttu-id="36d55-111">说明</span><span class="sxs-lookup"><span data-stu-id="36d55-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="36d55-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="36d55-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="36d55-113">否</span><span class="sxs-lookup"><span data-stu-id="36d55-113">No</span></span>   | <span data-ttu-id="36d55-114">定义外接程序访问关联的 API 或功能所需的扩展权限。</span><span class="sxs-lookup"><span data-stu-id="36d55-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="36d55-115">`ExtendedPermissions`示例</span><span class="sxs-lookup"><span data-stu-id="36d55-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="36d55-116">以下是`ExtendedPermissions`元素的示例。</span><span class="sxs-lookup"><span data-stu-id="36d55-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="36d55-117">包含于</span><span class="sxs-lookup"><span data-stu-id="36d55-117">Contained in</span></span>

[<span data-ttu-id="36d55-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="36d55-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="36d55-119">可以包含</span><span class="sxs-lookup"><span data-stu-id="36d55-119">Can contain</span></span>

[<span data-ttu-id="36d55-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="36d55-120">ExtendedPermission</span></span>](extendedpermission.md)
