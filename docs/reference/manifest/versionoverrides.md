# <a name="versionoverrides-element"></a><span data-ttu-id="95b91-101">VersionOverrides 元素</span><span class="sxs-lookup"><span data-stu-id="95b91-101">VersionOverrides element</span></span>

<span data-ttu-id="95b91-p101">根元素包含由加载项实现的加载项命令的信息。**VersionOverrides** 是清单中 [OfficeApp](./officeapp.md) 元素的子元素。此元素在清单架构 v1.1 及更高版本中受支持，但是在 VersionOverrides v1.0 或 v1.1 架构中进行定义。</span><span class="sxs-lookup"><span data-stu-id="95b91-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="95b91-105">属性</span><span class="sxs-lookup"><span data-stu-id="95b91-105">Attributes</span></span>

|  <span data-ttu-id="95b91-106">属性</span><span class="sxs-lookup"><span data-stu-id="95b91-106">Attribute</span></span>  |  <span data-ttu-id="95b91-107">必需</span><span class="sxs-lookup"><span data-stu-id="95b91-107">Required</span></span>  |  <span data-ttu-id="95b91-108">说明</span><span class="sxs-lookup"><span data-stu-id="95b91-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="95b91-109">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="95b91-109">**xmlns**</span></span>       |  <span data-ttu-id="95b91-110">是</span><span class="sxs-lookup"><span data-stu-id="95b91-110">Yes</span></span>  |  <span data-ttu-id="95b91-111">若 `xsi:type` 为 `VersionOverridesV1_0`，架构位置必须是 `http://schemas.microsoft.com/office/mailappversionoverrides`；若 `xsi:type` 为 `VersionOverridesV1_1`，架构位置必须是 `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`。</span><span class="sxs-lookup"><span data-stu-id="95b91-111">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="95b91-112">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="95b91-112">**xsi:type**</span></span>  |  <span data-ttu-id="95b91-113">是</span><span class="sxs-lookup"><span data-stu-id="95b91-113">Yes</span></span>  | <span data-ttu-id="95b91-p102">架构版本。目前的唯一有效值为 `VersionOverridesV1_0` 和 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="95b91-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="95b91-116">暂只有 Outlook 2016 支持 VersionOverrides v1.1 架构和 `VersionOverridesV1_1` 类型。</span><span class="sxs-lookup"><span data-stu-id="95b91-116">Note: Currently only Outlook 2016 supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="95b91-117">子元素</span><span class="sxs-lookup"><span data-stu-id="95b91-117">Child elements</span></span>

|  <span data-ttu-id="95b91-118">元素</span><span class="sxs-lookup"><span data-stu-id="95b91-118">Element</span></span> |  <span data-ttu-id="95b91-119">必需</span><span class="sxs-lookup"><span data-stu-id="95b91-119">Required</span></span>  |  <span data-ttu-id="95b91-120">说明</span><span class="sxs-lookup"><span data-stu-id="95b91-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="95b91-121">**说明**</span><span class="sxs-lookup"><span data-stu-id="95b91-121">**Description**</span></span>    |  <span data-ttu-id="95b91-122">No</span><span class="sxs-lookup"><span data-stu-id="95b91-122">No</span></span>   |  <span data-ttu-id="95b91-p103">描述加载项。这会替代清单中任何父级部分中的 `Description` 元素。说明文本包含在 [Rescources](./resources.md) 元素中的 **LongString** 元素的子元素中。**Description** 元素的 `resid` 属性被设置为包含文本的 `String` 元素的 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="95b91-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="95b91-127">**要求**</span><span class="sxs-lookup"><span data-stu-id="95b91-127">**Requirements**</span></span>  |  <span data-ttu-id="95b91-128">No</span><span class="sxs-lookup"><span data-stu-id="95b91-128">No</span></span>   |  <span data-ttu-id="95b91-p104">指定加载项要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。</span><span class="sxs-lookup"><span data-stu-id="95b91-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="95b91-131">Hosts</span><span class="sxs-lookup"><span data-stu-id="95b91-131">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="95b91-132">是</span><span class="sxs-lookup"><span data-stu-id="95b91-132">Yes</span></span>  |  <span data-ttu-id="95b91-p105">指定 Office 主机的集合。子级 Hosts 元素替代清单中父级部分中的 Hosts 元素。</span><span class="sxs-lookup"><span data-stu-id="95b91-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="95b91-135">资源</span><span class="sxs-lookup"><span data-stu-id="95b91-135">Resources</span></span>](./resources.md)    |  <span data-ttu-id="95b91-136">是</span><span class="sxs-lookup"><span data-stu-id="95b91-136">Yes</span></span>  | <span data-ttu-id="95b91-137">定义其他清单元素引用一组的资源（字符串、URL 和图像）。</span><span class="sxs-lookup"><span data-stu-id="95b91-137">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="95b91-138">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="95b91-138">**VersionOverrides**</span></span>    |  <span data-ttu-id="95b91-139">No</span><span class="sxs-lookup"><span data-stu-id="95b91-139">No</span></span>  | <span data-ttu-id="95b91-p106">在新版架构下定义加载项命令。有关详细信息，请参阅[实现多个版本](#implementing-multiple-versions)。</span><span class="sxs-lookup"><span data-stu-id="95b91-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="95b91-142">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="95b91-142">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="95b91-143">No</span><span class="sxs-lookup"><span data-stu-id="95b91-143">No</span></span>  | <span data-ttu-id="95b91-144">指定加载项关联 Web 应用程序的详细信息。</span><span class="sxs-lookup"><span data-stu-id="95b91-144">Specifies details about the add-in's associated Web application.</span></span> |



### <a name="versionoverrides-example"></a><span data-ttu-id="95b91-145">VersionOverrides 示例</span><span class="sxs-lookup"><span data-stu-id="95b91-145">VersionOverrides example</span></span>
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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="95b91-146">实现多个版本</span><span class="sxs-lookup"><span data-stu-id="95b91-146">Implementing multiple versions</span></span>

<span data-ttu-id="95b91-p107">清单可以实现 `VersionOverrides` 元素的多个版本，这些版本支持不同版本的 VersionOverrides 架构。为此，可以视情况支持新版架构中的新功能，同时仍支持不支持新功能的旧版客户端。</span><span class="sxs-lookup"><span data-stu-id="95b91-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="95b91-149">新版架构的 `VersionOverrides` 元素必须是旧版架构的 `VersionOverrides` 元素的子元素，才能实现多个版本。</span><span class="sxs-lookup"><span data-stu-id="95b91-149">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="95b91-150">子元素 `VersionOverrides`不会从父元素继承任何值。</span><span class="sxs-lookup"><span data-stu-id="95b91-150">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="95b91-151">若要实现 VersionOverrides v1.0 和 v1.1 架构，清单如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="95b91-151">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
...
</OfficeApp>
```
