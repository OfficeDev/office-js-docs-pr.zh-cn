---
title: 清单文件中的 VersionOverrides 元素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: ce65cdced1b3cf885cee09732c2cda0081a53cfc
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477878"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="5e0fc-102">VersionOverrides 元素</span><span class="sxs-lookup"><span data-stu-id="5e0fc-102">VersionOverrides element</span></span>

<span data-ttu-id="5e0fc-p101">根元素包含由外接程序实现的外接程序命令的信息。**VersionOverrides** 是清单中 [OfficeApp](./officeapp.md) 元素的子元素。此元素在清单架构 v1.1 及更高版本中受支持，但是在 VersionOverrides v1.0 或 v1.1 架构中进行定义。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="5e0fc-106">属性</span><span class="sxs-lookup"><span data-stu-id="5e0fc-106">Attributes</span></span>

|  <span data-ttu-id="5e0fc-107">属性</span><span class="sxs-lookup"><span data-stu-id="5e0fc-107">Attribute</span></span>  |  <span data-ttu-id="5e0fc-108">必需</span><span class="sxs-lookup"><span data-stu-id="5e0fc-108">Required</span></span>  |  <span data-ttu-id="5e0fc-109">说明</span><span class="sxs-lookup"><span data-stu-id="5e0fc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5e0fc-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="5e0fc-110">**xmlns**</span></span>       |  <span data-ttu-id="5e0fc-111">是</span><span class="sxs-lookup"><span data-stu-id="5e0fc-111">Yes</span></span>  |  <span data-ttu-id="5e0fc-112">若 `http://schemas.microsoft.com/office/mailappversionoverrides` 为 `xsi:type`，架构位置必须是 `VersionOverridesV1_0`；若 `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 为 `xsi:type`，架构位置必须是 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-112">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="5e0fc-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="5e0fc-113">**xsi:type**</span></span>  |  <span data-ttu-id="5e0fc-114">是</span><span class="sxs-lookup"><span data-stu-id="5e0fc-114">Yes</span></span>  | <span data-ttu-id="5e0fc-p102">架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="5e0fc-117">目前, 只有 Outlook 2016 或更高版本支持 VersionOverrides v1.1 架构和`VersionOverridesV1_1`类型。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-117">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5e0fc-118">子元素</span><span class="sxs-lookup"><span data-stu-id="5e0fc-118">Child elements</span></span>

|  <span data-ttu-id="5e0fc-119">元素</span><span class="sxs-lookup"><span data-stu-id="5e0fc-119">Element</span></span> |  <span data-ttu-id="5e0fc-120">必需</span><span class="sxs-lookup"><span data-stu-id="5e0fc-120">Required</span></span>  |  <span data-ttu-id="5e0fc-121">说明</span><span class="sxs-lookup"><span data-stu-id="5e0fc-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5e0fc-122">**说明**</span><span class="sxs-lookup"><span data-stu-id="5e0fc-122">**Description**</span></span>    |  <span data-ttu-id="5e0fc-123">否</span><span class="sxs-lookup"><span data-stu-id="5e0fc-123">No</span></span>   |  <span data-ttu-id="5e0fc-p103">描述外接程序。这会替代清单中任何父级部分中的 `Description` 元素。说明文本包含在 **Rescources** 元素中的 [LongString](./resources.md) 元素的子元素中。`resid` 元素的 \*\*\*\* 属性被设置为包含文本的 `id` 元素的 `String` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
| <span data-ttu-id="5e0fc-128">**EquivalentAddins**</span><span class="sxs-lookup"><span data-stu-id="5e0fc-128">**EquivalentAddins**</span></span> | <span data-ttu-id="5e0fc-129">否</span><span class="sxs-lookup"><span data-stu-id="5e0fc-129">No</span></span> | <span data-ttu-id="5e0fc-130">指定与等效的 COM 外接程序、XLL 或这两者的向后兼容性。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-130">Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.</span></span> |
|  <span data-ttu-id="5e0fc-131">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="5e0fc-131">**Requirements**</span></span>  |  <span data-ttu-id="5e0fc-132">否</span><span class="sxs-lookup"><span data-stu-id="5e0fc-132">No</span></span>   |  <span data-ttu-id="5e0fc-p104">指定外接程序要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="5e0fc-135">Hosts</span><span class="sxs-lookup"><span data-stu-id="5e0fc-135">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="5e0fc-136">是</span><span class="sxs-lookup"><span data-stu-id="5e0fc-136">Yes</span></span>  |  <span data-ttu-id="5e0fc-p105">指定 Office 主机的集合。子级 Hosts 元素替代清单中父级部分中的 Hosts 元素。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="5e0fc-139">Resources</span><span class="sxs-lookup"><span data-stu-id="5e0fc-139">Resources</span></span>](./resources.md)    |  <span data-ttu-id="5e0fc-140">是</span><span class="sxs-lookup"><span data-stu-id="5e0fc-140">Yes</span></span>  | <span data-ttu-id="5e0fc-141">定义其他清单元素引用的资源集合（字符串、URL 和图像）。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-141">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="5e0fc-142">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="5e0fc-142">EquivalentAddins</span></span>](./equivalentaddins.md)    |  <span data-ttu-id="5e0fc-143">否</span><span class="sxs-lookup"><span data-stu-id="5e0fc-143">No</span></span>  | <span data-ttu-id="5e0fc-144">指定与 web 外接程序等效的本机 (COM/XLL) 加载项。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-144">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="5e0fc-145">如果安装了等效的本机加载项, 则不会激活 web 外接程序。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-145">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="5e0fc-146">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="5e0fc-146">**VersionOverrides**</span></span>    |  <span data-ttu-id="5e0fc-147">否</span><span class="sxs-lookup"><span data-stu-id="5e0fc-147">No</span></span>  | <span data-ttu-id="5e0fc-p107">在新版架构下定义外接程序命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-p107">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="5e0fc-150">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="5e0fc-150">WebApplicationInfo</span></span>](./webapplicationinfo.md)    |  <span data-ttu-id="5e0fc-151">否</span><span class="sxs-lookup"><span data-stu-id="5e0fc-151">No</span></span>  | <span data-ttu-id="5e0fc-152">指定有关使用安全令牌颁发者 (如 Azure Active Directory v2.0) 的加载项注册的详细信息。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-152">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="5e0fc-153">VersionOverrides 示例</span><span class="sxs-lookup"><span data-stu-id="5e0fc-153">VersionOverrides example</span></span>

<span data-ttu-id="5e0fc-154">下面是典型`<VersionOverrides>`元素的一个示例, 其中包括一些不需要但通常使用的子元素。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-154">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp>
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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="5e0fc-155">实现多个版本</span><span class="sxs-lookup"><span data-stu-id="5e0fc-155">Implementing multiple versions</span></span>

<span data-ttu-id="5e0fc-p108">清单可以实现 `VersionOverrides` 元素的多个版本，这些版本支持不同版本的 VersionOverrides 架构。为此，可以视情况支持新版架构中的新功能，同时仍支持不支持新功能的旧版客户端。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-p108">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="5e0fc-158">新版架构的 `VersionOverrides` 元素必须是旧版架构的 `VersionOverrides` 元素的子元素，才能实现多个版本。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-158">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="5e0fc-159">`VersionOverrides` 子元素不会从父元素继承任何值。</span><span class="sxs-lookup"><span data-stu-id="5e0fc-159">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="5e0fc-160">若要实现 VersionOverrides v1.0 和 v1.1 架构，清单如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="5e0fc-160">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
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
