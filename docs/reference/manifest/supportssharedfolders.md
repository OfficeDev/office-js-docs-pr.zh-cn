---
title: 清单文件中的 SupportsSharedFolders 元素
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: e76d17b618e2aaf15724f15ee6695a932172bba3
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325225"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="a9365-102">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="a9365-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="a9365-103">定义 Outlook 加载项在代理应用场景中是否可用。</span><span class="sxs-lookup"><span data-stu-id="a9365-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="a9365-104">**SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="a9365-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="a9365-105">默认情况下，此元素设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="a9365-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a9365-106">只有 Outlook 网页和 Windows 支持**SupportsSharedFolders**元素。</span><span class="sxs-lookup"><span data-stu-id="a9365-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="a9365-107">对此元素的支持是在要求集1.8 中引入的。</span><span class="sxs-lookup"><span data-stu-id="a9365-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="a9365-108">请查看支持此要求集的[客户端和平台](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="a9365-108">See [clients and platforms](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="a9365-109">下面是**SupportsSharedFolders**元素的一个示例。</span><span class="sxs-lookup"><span data-stu-id="a9365-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
