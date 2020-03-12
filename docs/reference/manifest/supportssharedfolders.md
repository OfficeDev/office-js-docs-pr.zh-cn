---
title: 清单文件中的 SupportsSharedFolders 元素
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 81401b79f4c443305e376df7a66a07d916393d17
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596751"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="edcbd-102">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="edcbd-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="edcbd-103">定义 Outlook 加载项在代理应用场景中是否可用。</span><span class="sxs-lookup"><span data-stu-id="edcbd-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="edcbd-104">**SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="edcbd-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="edcbd-105">默认情况下，此元素设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="edcbd-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="edcbd-106">只有 Outlook 网页和 Windows 支持**SupportsSharedFolders**元素。</span><span class="sxs-lookup"><span data-stu-id="edcbd-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="edcbd-107">对此元素的支持是在要求集1.8 中引入的。</span><span class="sxs-lookup"><span data-stu-id="edcbd-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="edcbd-108">请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="edcbd-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="edcbd-109">下面是**SupportsSharedFolders**元素的一个示例。</span><span class="sxs-lookup"><span data-stu-id="edcbd-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
  </VersionOverrides>
</VersionOverrides>
...
```
