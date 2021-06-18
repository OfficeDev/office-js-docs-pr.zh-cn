---
title: 清单文件中的 SupportsSharedFolders 元素
description: SupportsSharedFolders 元素定义 Outlook外接程序在共享文件夹和共享邮箱方案中是否可用。
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 43f2c60664a6822b714023246cfa044e179e9a55
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007781"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="fb06f-103">SupportsSharedFolders 元素</span><span class="sxs-lookup"><span data-stu-id="fb06f-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="fb06f-104">定义 Outlook 外接程序现在在预览版 (中是否可用) 以及共享文件夹 (即委派访问权限) 方案。</span><span class="sxs-lookup"><span data-stu-id="fb06f-104">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="fb06f-105">**SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="fb06f-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="fb06f-106">默认情况下，此元素设置为 *false*。</span><span class="sxs-lookup"><span data-stu-id="fb06f-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fb06f-107">要求集 1.8 中引入了对此元素的支持。</span><span class="sxs-lookup"><span data-stu-id="fb06f-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="fb06f-108">请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="fb06f-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="fb06f-109">下面是 **SupportsSharedFolders 元素的一** 个示例。</span><span class="sxs-lookup"><span data-stu-id="fb06f-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
