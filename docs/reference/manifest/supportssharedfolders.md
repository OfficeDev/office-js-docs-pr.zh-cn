---
title: 清单文件中的 SupportsSharedFolders 元素
description: SupportsSharedFolders 元素定义 Outlook 加载项在委托方案中是否可用。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 786a4763450d78cb16c9baafc81701758af54787
ms.sourcegitcommit: 6fa29989dfaec4dfa0f8df3fe5fb038d7afbae30
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/16/2020
ms.locfileid: "48487878"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="db8f9-103">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="db8f9-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="db8f9-104">定义 Outlook 加载项在代理应用场景中是否可用。</span><span class="sxs-lookup"><span data-stu-id="db8f9-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="db8f9-105">**SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="db8f9-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="db8f9-106">默认情况下，此元素设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="db8f9-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="db8f9-107">对此元素的支持是在要求集1.8 中引入的。</span><span class="sxs-lookup"><span data-stu-id="db8f9-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="db8f9-108">请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="db8f9-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="db8f9-109">下面是 **SupportsSharedFolders** 元素的一个示例。</span><span class="sxs-lookup"><span data-stu-id="db8f9-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
