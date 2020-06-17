---
title: 清单文件中的 SupportsSharedFolders 元素
description: SupportsSharedFolders 元素定义 Outlook 加载项在委托方案中是否可用。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 3835f7060cc52a72ff0a5ed4dbdb9f1e09258669
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608710"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="b13a0-103">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="b13a0-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="b13a0-104">定义 Outlook 加载项在代理应用场景中是否可用。</span><span class="sxs-lookup"><span data-stu-id="b13a0-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="b13a0-105">**SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="b13a0-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="b13a0-106">默认情况下，此元素设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="b13a0-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b13a0-107">只有 Outlook 网页和 Windows 支持**SupportsSharedFolders**元素。</span><span class="sxs-lookup"><span data-stu-id="b13a0-107">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="b13a0-108">对此元素的支持是在要求集1.8 中引入的。</span><span class="sxs-lookup"><span data-stu-id="b13a0-108">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="b13a0-109">请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="b13a0-109">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="b13a0-110">下面是**SupportsSharedFolders**元素的一个示例。</span><span class="sxs-lookup"><span data-stu-id="b13a0-110">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
