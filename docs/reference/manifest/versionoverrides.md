---
title: 清单文件中的 VersionOverrides 元素
description: Office 外接程序清单的 VersionOverrides 元素参考文档 (XML) 文件。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 772eaa416909d24f8035ed3e1445d1e4f06a244e
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771303"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="a62be-103">VersionOverrides 元素</span><span class="sxs-lookup"><span data-stu-id="a62be-103">VersionOverrides element</span></span>

<span data-ttu-id="a62be-p101">根元素包含由外接程序实现的外接程序命令的信息。**VersionOverrides** 是清单中 [OfficeApp](officeapp.md) 元素的子元素。此元素在清单架构 v1.1 及更高版本中受支持，但是在 VersionOverrides v1.0 或 v1.1 架构中进行定义。</span><span class="sxs-lookup"><span data-stu-id="a62be-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="a62be-107">属性</span><span class="sxs-lookup"><span data-stu-id="a62be-107">Attributes</span></span>

|  <span data-ttu-id="a62be-108">属性</span><span class="sxs-lookup"><span data-stu-id="a62be-108">Attribute</span></span>  |  <span data-ttu-id="a62be-109">必需</span><span class="sxs-lookup"><span data-stu-id="a62be-109">Required</span></span>  |  <span data-ttu-id="a62be-110">说明</span><span class="sxs-lookup"><span data-stu-id="a62be-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a62be-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="a62be-111">**xmlns**</span></span>       |  <span data-ttu-id="a62be-112">是</span><span class="sxs-lookup"><span data-stu-id="a62be-112">Yes</span></span>  |  <span data-ttu-id="a62be-113">VersionOverrides 架构命名空间。</span><span class="sxs-lookup"><span data-stu-id="a62be-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="a62be-114">允许的值因此 `<VersionOverrides>` 元素 **的 xsi：type** 值和 **父元素的 xsi：type** `<OfficeApp>` 值而异。</span><span class="sxs-lookup"><span data-stu-id="a62be-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="a62be-115">请参阅 [下面的命名空间](#namespace-values) 值。</span><span class="sxs-lookup"><span data-stu-id="a62be-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="a62be-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="a62be-116">**xsi:type**</span></span>  |  <span data-ttu-id="a62be-117">是</span><span class="sxs-lookup"><span data-stu-id="a62be-117">Yes</span></span>  | <span data-ttu-id="a62be-p103">架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="a62be-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="a62be-120">命名空间值</span><span class="sxs-lookup"><span data-stu-id="a62be-120">Namespace values</span></span>

<span data-ttu-id="a62be-121">下面列出了 **xmlns** 值的必需值，具体取决于 **父元素的 xsi：type** `<OfficeApp>` 值。</span><span class="sxs-lookup"><span data-stu-id="a62be-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="a62be-122">**TaskPaneApp** 仅支持 VersionOverrides 版本 1.0，xmlns 应为 `http://schemas.microsoft.com/office/taskpaneappversionoverrides` 。</span><span class="sxs-lookup"><span data-stu-id="a62be-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="a62be-123">**ContentApp** 仅支持 VersionOverrides 版本 1.0，xmlns 应为 `http://schemas.microsoft.com/office/contentappversionoverrides` 。</span><span class="sxs-lookup"><span data-stu-id="a62be-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="a62be-124">**MailApp** 支持 VersionOverrides 版本 1.0 和 1.1，因此 **xmlns** 的值因此元素 `<VersionOverrides>` 的 **xsi：type** 值而异：</span><span class="sxs-lookup"><span data-stu-id="a62be-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="a62be-125">当 **xsi：type** 为 `VersionOverridesV1_0` 时 **，xmlns** 必须为 `http://schemas.microsoft.com/office/mailappversionoverrides` 。</span><span class="sxs-lookup"><span data-stu-id="a62be-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="a62be-126">当 **xsi：type** 为 `VersionOverridesV1_1` 时 **，xmlns** 必须为 `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。</span><span class="sxs-lookup"><span data-stu-id="a62be-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="a62be-127">目前，只有 Outlook 2016 或更高版本支持 VersionOverrides v1.1 架构和 `VersionOverridesV1_1` 类型。</span><span class="sxs-lookup"><span data-stu-id="a62be-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a62be-128">子元素</span><span class="sxs-lookup"><span data-stu-id="a62be-128">Child elements</span></span>

|  <span data-ttu-id="a62be-129">元素</span><span class="sxs-lookup"><span data-stu-id="a62be-129">Element</span></span> |  <span data-ttu-id="a62be-130">必需</span><span class="sxs-lookup"><span data-stu-id="a62be-130">Required</span></span>  |  <span data-ttu-id="a62be-131">说明</span><span class="sxs-lookup"><span data-stu-id="a62be-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a62be-132">**说明**</span><span class="sxs-lookup"><span data-stu-id="a62be-132">**Description**</span></span>    |  <span data-ttu-id="a62be-133">否</span><span class="sxs-lookup"><span data-stu-id="a62be-133">No</span></span>   |  <span data-ttu-id="a62be-134">描述外接程序。</span><span class="sxs-lookup"><span data-stu-id="a62be-134">Describes the add-in.</span></span> <span data-ttu-id="a62be-135">这会替代清单中任何父级部分中的 `Description` 元素。</span><span class="sxs-lookup"><span data-stu-id="a62be-135">This overrides the `Description` element in any parent portion of the manifest.</span></span> <span data-ttu-id="a62be-136">说明文本包含在 **Rescources** 元素中的 [LongString](resources.md) 元素的子元素中。</span><span class="sxs-lookup"><span data-stu-id="a62be-136">The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element.</span></span> <span data-ttu-id="a62be-137">`resid` **Description** 元素的属性不能超过 32 个字符，并设置为包含文本 `id` `String` 的元素的属性值。</span><span class="sxs-lookup"><span data-stu-id="a62be-137">The `resid` attribute of the **Description** element can be no more than 32 characters and is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="a62be-138">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="a62be-138">**Requirements**</span></span>  |  <span data-ttu-id="a62be-139">否</span><span class="sxs-lookup"><span data-stu-id="a62be-139">No</span></span>   |  <span data-ttu-id="a62be-p105">指定外接程序要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。</span><span class="sxs-lookup"><span data-stu-id="a62be-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="a62be-142">Hosts</span><span class="sxs-lookup"><span data-stu-id="a62be-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="a62be-143">是</span><span class="sxs-lookup"><span data-stu-id="a62be-143">Yes</span></span>  |  <span data-ttu-id="a62be-144">指定 Office 应用程序的集合。</span><span class="sxs-lookup"><span data-stu-id="a62be-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="a62be-145">子 Hosts 元素替代清单的父部分中的 Hosts 元素。</span><span class="sxs-lookup"><span data-stu-id="a62be-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="a62be-146">Resources</span><span class="sxs-lookup"><span data-stu-id="a62be-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="a62be-147">是</span><span class="sxs-lookup"><span data-stu-id="a62be-147">Yes</span></span>  | <span data-ttu-id="a62be-148">定义其他清单元素引用的资源集合（字符串、URL 和图像）。</span><span class="sxs-lookup"><span data-stu-id="a62be-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="a62be-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="a62be-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="a62be-150">否</span><span class="sxs-lookup"><span data-stu-id="a62be-150">No</span></span>  | <span data-ttu-id="a62be-151">指定与 web (等效) COM/XLL 加载项的本机属性。</span><span class="sxs-lookup"><span data-stu-id="a62be-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="a62be-152">如果安装了等效的本机加载项，则不激活 Web 加载项。</span><span class="sxs-lookup"><span data-stu-id="a62be-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="a62be-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="a62be-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="a62be-154">否</span><span class="sxs-lookup"><span data-stu-id="a62be-154">No</span></span>  | <span data-ttu-id="a62be-p108">在新版架构下定义外接程序命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。</span><span class="sxs-lookup"><span data-stu-id="a62be-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="a62be-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="a62be-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="a62be-158">否</span><span class="sxs-lookup"><span data-stu-id="a62be-158">No</span></span>  | <span data-ttu-id="a62be-159">指定有关向安全令牌颁发者（如 Azure Active Directory V2.0）注册加载项的详细信息。</span><span class="sxs-lookup"><span data-stu-id="a62be-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="a62be-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="a62be-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="a62be-161">否</span><span class="sxs-lookup"><span data-stu-id="a62be-161">No</span></span>  |  <span data-ttu-id="a62be-162">指定扩展权限的集合。</span><span class="sxs-lookup"><span data-stu-id="a62be-162">Specifies a collection of extended permissions.</span></span><br><br><span data-ttu-id="a62be-163">**重要** 提示：由于 [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API 当前处于预览阶段，因此使用元素的加载项无法发布到 AppSource 或通过集中部署 `ExtendedPermissions` 进行部署。</span><span class="sxs-lookup"><span data-stu-id="a62be-163">**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="a62be-164">VersionOverrides 示例</span><span class="sxs-lookup"><span data-stu-id="a62be-164">VersionOverrides example</span></span>

<span data-ttu-id="a62be-165">下面是一个典型元素的示例，其中包括一些不需要但 `<VersionOverrides>` 通常使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="a62be-165">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="a62be-166">实现多个版本</span><span class="sxs-lookup"><span data-stu-id="a62be-166">Implementing multiple versions</span></span>

<span data-ttu-id="a62be-p109">清单可以实现 `VersionOverrides` 元素的多个版本，这些版本支持不同版本的 VersionOverrides 架构。为此，可以视情况支持新版架构中的新功能，同时仍支持不支持新功能的旧版客户端。</span><span class="sxs-lookup"><span data-stu-id="a62be-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="a62be-169">新版架构的 `VersionOverrides` 元素必须是旧版架构的 `VersionOverrides` 元素的子元素，才能实现多个版本。</span><span class="sxs-lookup"><span data-stu-id="a62be-169">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="a62be-170">`VersionOverrides` 子元素不会从父元素继承任何值。</span><span class="sxs-lookup"><span data-stu-id="a62be-170">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="a62be-171">若要实现 VersionOverrides v1.0 和 v1.1 架构，清单如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="a62be-171">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
