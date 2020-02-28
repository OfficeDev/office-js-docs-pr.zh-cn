---
title: 清单文件中的 VersionOverrides 元素
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 0afa3183e34a736a878217c079b7b8d0259be5b1
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324797"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="93685-102">VersionOverrides 元素</span><span class="sxs-lookup"><span data-stu-id="93685-102">VersionOverrides element</span></span>

<span data-ttu-id="93685-p101">根元素包含由外接程序实现的外接程序命令的信息。**VersionOverrides** 是清单中 [OfficeApp](./officeapp.md) 元素的子元素。此元素在清单架构 v1.1 及更高版本中受支持，但是在 VersionOverrides v1.0 或 v1.1 架构中进行定义。</span><span class="sxs-lookup"><span data-stu-id="93685-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="93685-106">属性</span><span class="sxs-lookup"><span data-stu-id="93685-106">Attributes</span></span>

|  <span data-ttu-id="93685-107">属性</span><span class="sxs-lookup"><span data-stu-id="93685-107">Attribute</span></span>  |  <span data-ttu-id="93685-108">必需</span><span class="sxs-lookup"><span data-stu-id="93685-108">Required</span></span>  |  <span data-ttu-id="93685-109">说明</span><span class="sxs-lookup"><span data-stu-id="93685-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="93685-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="93685-110">**xmlns**</span></span>       |  <span data-ttu-id="93685-111">是</span><span class="sxs-lookup"><span data-stu-id="93685-111">Yes</span></span>  |  <span data-ttu-id="93685-112">VersionOverrides 架构命名空间。</span><span class="sxs-lookup"><span data-stu-id="93685-112">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="93685-113">根据此`<VersionOverrides>`元素的**xsi： type**值和父`<OfficeApp>`元素的**xsi： type**值，允许的值会有所不同。</span><span class="sxs-lookup"><span data-stu-id="93685-113">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="93685-114">请参阅下面的[命名空间值](#namespace-values)。</span><span class="sxs-lookup"><span data-stu-id="93685-114">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="93685-115">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="93685-115">**xsi:type**</span></span>  |  <span data-ttu-id="93685-116">是</span><span class="sxs-lookup"><span data-stu-id="93685-116">Yes</span></span>  | <span data-ttu-id="93685-p103">架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="93685-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="93685-119">命名空间值</span><span class="sxs-lookup"><span data-stu-id="93685-119">Namespace values</span></span>

<span data-ttu-id="93685-120">下面列出了**xmlns**值所需的值，具体取决于父`<OfficeApp>`元素的**xsi： type**值。</span><span class="sxs-lookup"><span data-stu-id="93685-120">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="93685-121">**TaskPaneApp**仅支持 VersionOverrides 的1.0 版，而**xmlns**应为`http://schemas.microsoft.com/office/taskpaneappversionoverrides`。</span><span class="sxs-lookup"><span data-stu-id="93685-121">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="93685-122">**ContentApp**仅支持 VersionOverrides 的1.0 版，而**xmlns**应为`http://schemas.microsoft.com/office/contentappversionoverrides`。</span><span class="sxs-lookup"><span data-stu-id="93685-122">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="93685-123">**MailApp**支持 VersionOverrides 的版本1.0 和1.1，因此根据`<VersionOverrides>`此元素的**xsi： type**值， **xmlns**的值会有所不同：</span><span class="sxs-lookup"><span data-stu-id="93685-123">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="93685-124">当**xsi： type**为`VersionOverridesV1_0`时， **xmlns**必须`http://schemas.microsoft.com/office/mailappversionoverrides`为。</span><span class="sxs-lookup"><span data-stu-id="93685-124">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="93685-125">当**xsi： type**为`VersionOverridesV1_1`时， **xmlns**必须`http://schemas.microsoft.com/office/mailappversionoverrides/1.1`为。</span><span class="sxs-lookup"><span data-stu-id="93685-125">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="93685-126">目前，只有 Outlook 2016 或更高版本支持 VersionOverrides v1.1 架构和`VersionOverridesV1_1`类型。</span><span class="sxs-lookup"><span data-stu-id="93685-126">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="93685-127">子元素</span><span class="sxs-lookup"><span data-stu-id="93685-127">Child elements</span></span>

|  <span data-ttu-id="93685-128">元素</span><span class="sxs-lookup"><span data-stu-id="93685-128">Element</span></span> |  <span data-ttu-id="93685-129">必需</span><span class="sxs-lookup"><span data-stu-id="93685-129">Required</span></span>  |  <span data-ttu-id="93685-130">说明</span><span class="sxs-lookup"><span data-stu-id="93685-130">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="93685-131">**说明**</span><span class="sxs-lookup"><span data-stu-id="93685-131">**Description**</span></span>    |  <span data-ttu-id="93685-132">否</span><span class="sxs-lookup"><span data-stu-id="93685-132">No</span></span>   |  <span data-ttu-id="93685-p104">描述外接程序。这会替代清单中任何父级部分中的 `Description` 元素。说明文本包含在 **Rescources** 元素中的 [LongString](./resources.md) 元素的子元素中。`resid` 元素的 \*\*\*\* 属性被设置为包含文本的 `id` 元素的 `String` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="93685-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="93685-137">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="93685-137">**Requirements**</span></span>  |  <span data-ttu-id="93685-138">否</span><span class="sxs-lookup"><span data-stu-id="93685-138">No</span></span>   |  <span data-ttu-id="93685-p105">指定外接程序要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。</span><span class="sxs-lookup"><span data-stu-id="93685-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="93685-141">Hosts</span><span class="sxs-lookup"><span data-stu-id="93685-141">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="93685-142">是</span><span class="sxs-lookup"><span data-stu-id="93685-142">Yes</span></span>  |  <span data-ttu-id="93685-p106">指定 Office 主机的集合。子级 Hosts 元素替代清单中父级部分中的 Hosts 元素。</span><span class="sxs-lookup"><span data-stu-id="93685-p106">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="93685-145">Resources</span><span class="sxs-lookup"><span data-stu-id="93685-145">Resources</span></span>](./resources.md)    |  <span data-ttu-id="93685-146">是</span><span class="sxs-lookup"><span data-stu-id="93685-146">Yes</span></span>  | <span data-ttu-id="93685-147">定义其他清单元素引用的资源集合（字符串、URL 和图像）。</span><span class="sxs-lookup"><span data-stu-id="93685-147">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="93685-148">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="93685-148">EquivalentAddins</span></span>](./equivalentaddins.md)    |  <span data-ttu-id="93685-149">否</span><span class="sxs-lookup"><span data-stu-id="93685-149">No</span></span>  | <span data-ttu-id="93685-150">指定与 web 外接程序等效的本机（COM/XLL）加载项。</span><span class="sxs-lookup"><span data-stu-id="93685-150">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="93685-151">如果安装了等效的本机加载项，则不会激活 web 外接程序。</span><span class="sxs-lookup"><span data-stu-id="93685-151">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="93685-152">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="93685-152">**VersionOverrides**</span></span>    |  <span data-ttu-id="93685-153">否</span><span class="sxs-lookup"><span data-stu-id="93685-153">No</span></span>  | <span data-ttu-id="93685-p108">在新版架构下定义外接程序命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。</span><span class="sxs-lookup"><span data-stu-id="93685-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="93685-156">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="93685-156">WebApplicationInfo</span></span>](./webapplicationinfo.md)    |  <span data-ttu-id="93685-157">否</span><span class="sxs-lookup"><span data-stu-id="93685-157">No</span></span>  | <span data-ttu-id="93685-158">指定有关使用安全令牌颁发者（如 Azure Active Directory v2.0）的加载项注册的详细信息。</span><span class="sxs-lookup"><span data-stu-id="93685-158">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="93685-159">VersionOverrides 示例</span><span class="sxs-lookup"><span data-stu-id="93685-159">VersionOverrides example</span></span>

<span data-ttu-id="93685-160">下面是典型`<VersionOverrides>`元素的一个示例，其中包括一些不需要但通常使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="93685-160">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="93685-161">实现多个版本</span><span class="sxs-lookup"><span data-stu-id="93685-161">Implementing multiple versions</span></span>

<span data-ttu-id="93685-p109">清单可以实现 `VersionOverrides` 元素的多个版本，这些版本支持不同版本的 VersionOverrides 架构。为此，可以视情况支持新版架构中的新功能，同时仍支持不支持新功能的旧版客户端。</span><span class="sxs-lookup"><span data-stu-id="93685-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="93685-164">新版架构的 `VersionOverrides` 元素必须是旧版架构的 `VersionOverrides` 元素的子元素，才能实现多个版本。</span><span class="sxs-lookup"><span data-stu-id="93685-164">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="93685-165">`VersionOverrides` 子元素不会从父元素继承任何值。</span><span class="sxs-lookup"><span data-stu-id="93685-165">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="93685-166">若要实现 VersionOverrides v1.0 和 v1.1 架构，清单如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="93685-166">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
