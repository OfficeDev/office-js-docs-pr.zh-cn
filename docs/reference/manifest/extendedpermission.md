---
title: 清单文件中的 ExtendedPermission 元素
description: 定义外接程序访问关联的 API 或功能所需的扩展权限。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 996cac59c44220d05165c7be6ae7c3d79d853271
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626398"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="1febc-103">`ExtendedPermission` 网元</span><span class="sxs-lookup"><span data-stu-id="1febc-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="1febc-104">定义外接程序访问关联的 API 或功能所需的扩展权限。</span><span class="sxs-lookup"><span data-stu-id="1febc-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="1febc-105">`ExtendedPermission`元素是[ExtendedPermissions](extendedpermissions.md)的子元素。</span><span class="sxs-lookup"><span data-stu-id="1febc-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1febc-106">对此元素的支持是在要求集1.9 中引入的。</span><span class="sxs-lookup"><span data-stu-id="1febc-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="1febc-107">请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="1febc-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="1febc-108">可用扩展权限</span><span class="sxs-lookup"><span data-stu-id="1febc-108">Available extended permissions</span></span>

<span data-ttu-id="1febc-109">以下是可用的值。</span><span class="sxs-lookup"><span data-stu-id="1febc-109">The following are the available values.</span></span>

|<span data-ttu-id="1febc-110">可用值</span><span class="sxs-lookup"><span data-stu-id="1febc-110">Available value</span></span>|<span data-ttu-id="1febc-111">说明</span><span class="sxs-lookup"><span data-stu-id="1febc-111">Description</span></span>|<span data-ttu-id="1febc-112">Hosts</span><span class="sxs-lookup"><span data-stu-id="1febc-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="1febc-113">声明外接程序使用的是 [appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API。</span><span class="sxs-lookup"><span data-stu-id="1febc-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="1febc-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="1febc-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="1febc-115">`ExtendedPermission` 示例</span><span class="sxs-lookup"><span data-stu-id="1febc-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="1febc-116">以下是元素的示例 `ExtendedPermission` 。</span><span class="sxs-lookup"><span data-stu-id="1febc-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="1febc-117">包含于</span><span class="sxs-lookup"><span data-stu-id="1febc-117">Contained in</span></span>

[<span data-ttu-id="1febc-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="1febc-118">ExtendedPermissions</span></span>](extendedpermissions.md)
