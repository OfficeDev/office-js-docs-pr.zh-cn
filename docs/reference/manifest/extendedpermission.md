---
title: 清单文件中的 ExtendedPermission 元素
description: 定义外接程序访问关联的 API 或功能所需的扩展权限。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: ca4c2d12cb9a5be159c22712b631d0bde42e48ed
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611538"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="25fde-103">`ExtendedPermission`网元</span><span class="sxs-lookup"><span data-stu-id="25fde-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="25fde-104">定义外接程序访问关联的 API 或功能所需的扩展权限。</span><span class="sxs-lookup"><span data-stu-id="25fde-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="25fde-105">`ExtendedPermission`元素是[ExtendedPermissions](extendedpermissions.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="25fde-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="25fde-106">此元素仅适用于针对 Exchange Online 的[Outlook 外接程序预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。</span><span class="sxs-lookup"><span data-stu-id="25fde-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="25fde-107">使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。</span><span class="sxs-lookup"><span data-stu-id="25fde-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="25fde-108">可用扩展权限</span><span class="sxs-lookup"><span data-stu-id="25fde-108">Available extended permissions</span></span>

<span data-ttu-id="25fde-109">以下是可用的值。</span><span class="sxs-lookup"><span data-stu-id="25fde-109">The following are the available values.</span></span>

|<span data-ttu-id="25fde-110">可用值</span><span class="sxs-lookup"><span data-stu-id="25fde-110">Available value</span></span>|<span data-ttu-id="25fde-111">Description</span><span class="sxs-lookup"><span data-stu-id="25fde-111">Description</span></span>|<span data-ttu-id="25fde-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="25fde-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="25fde-113">声明外接程序使用的是[appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API。</span><span class="sxs-lookup"><span data-stu-id="25fde-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="25fde-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="25fde-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="25fde-115">`ExtendedPermission`示例</span><span class="sxs-lookup"><span data-stu-id="25fde-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="25fde-116">以下是元素的示例 `ExtendedPermission` 。</span><span class="sxs-lookup"><span data-stu-id="25fde-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="25fde-117">包含于</span><span class="sxs-lookup"><span data-stu-id="25fde-117">Contained in</span></span>

[<span data-ttu-id="25fde-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="25fde-118">ExtendedPermissions</span></span>](extendedpermissions.md)
