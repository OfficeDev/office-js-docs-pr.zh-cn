# <a name="host-element"></a><span data-ttu-id="fd7fc-101">Host 元素</span><span class="sxs-lookup"><span data-stu-id="fd7fc-101">Host element</span></span>

<span data-ttu-id="fd7fc-102">指定应在其中激活外接程序的单个 Office 应用程序类型。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="fd7fc-103">**Host** 元素的语法根据该元素是否在[基本清单](#basic-manifest)中或 [VersionOverrides](#versionoverrides-node) 节点中定义而不同。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-103">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="fd7fc-104">但功能相同。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="fd7fc-105">基本清单</span><span class="sxs-lookup"><span data-stu-id="fd7fc-105">Basic manifest</span></span>

<span data-ttu-id="fd7fc-106">在基本清单（在 [OfficeApp](officeapp.md) 下）中定义时，主机类型由 `Name` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="fd7fc-107">属性</span><span class="sxs-lookup"><span data-stu-id="fd7fc-107">Attributes</span></span>

| <span data-ttu-id="fd7fc-108">属性</span><span class="sxs-lookup"><span data-stu-id="fd7fc-108">Attribute</span></span>     | <span data-ttu-id="fd7fc-109">类型</span><span class="sxs-lookup"><span data-stu-id="fd7fc-109">Type</span></span>   | <span data-ttu-id="fd7fc-110">是否必需</span><span class="sxs-lookup"><span data-stu-id="fd7fc-110">Required</span></span> | <span data-ttu-id="fd7fc-111">说明</span><span class="sxs-lookup"><span data-stu-id="fd7fc-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="fd7fc-112">名称</span><span class="sxs-lookup"><span data-stu-id="fd7fc-112">Name</span></span>](#name) | <span data-ttu-id="fd7fc-113">字符串</span><span class="sxs-lookup"><span data-stu-id="fd7fc-113">string</span></span> | <span data-ttu-id="fd7fc-114">必需</span><span class="sxs-lookup"><span data-stu-id="fd7fc-114">required</span></span> | <span data-ttu-id="fd7fc-115">Office 主机应用程序的类型名称。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="fd7fc-116">名称</span><span class="sxs-lookup"><span data-stu-id="fd7fc-116">Name</span></span>
<span data-ttu-id="fd7fc-p102">指定此外接程序面向的主机类型。值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="fd7fc-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="fd7fc-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-119">`Document` (Word)</span></span>
- <span data-ttu-id="fd7fc-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-120">`Database` (Access)</span></span>
- <span data-ttu-id="fd7fc-121">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="fd7fc-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="fd7fc-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="fd7fc-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-124">`Project` (Project)</span></span>
- <span data-ttu-id="fd7fc-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="fd7fc-126">示例</span><span class="sxs-lookup"><span data-stu-id="fd7fc-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="fd7fc-127">VersionOverrides 节点</span><span class="sxs-lookup"><span data-stu-id="fd7fc-127">VersionOverrides node</span></span>
<span data-ttu-id="fd7fc-128">在 [VersionOverrides](versionoverrides.md) 中定义时，主机类型由 `xsi:type` 属性决定。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="fd7fc-129">属性</span><span class="sxs-lookup"><span data-stu-id="fd7fc-129">Attributes</span></span>

|  <span data-ttu-id="fd7fc-130">属性</span><span class="sxs-lookup"><span data-stu-id="fd7fc-130">Attribute</span></span>  |  <span data-ttu-id="fd7fc-131">是否必需</span><span class="sxs-lookup"><span data-stu-id="fd7fc-131">Required</span></span>  |  <span data-ttu-id="fd7fc-132">说明</span><span class="sxs-lookup"><span data-stu-id="fd7fc-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fd7fc-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fd7fc-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="fd7fc-134">是</span><span class="sxs-lookup"><span data-stu-id="fd7fc-134">Yes</span></span>  | <span data-ttu-id="fd7fc-135">描述这些设置适用的 Office 主机。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="fd7fc-136">子元素</span><span class="sxs-lookup"><span data-stu-id="fd7fc-136">Child elements</span></span>

|  <span data-ttu-id="fd7fc-137">元素</span><span class="sxs-lookup"><span data-stu-id="fd7fc-137">Element</span></span> |  <span data-ttu-id="fd7fc-138">是否必需</span><span class="sxs-lookup"><span data-stu-id="fd7fc-138">Required</span></span>  |  <span data-ttu-id="fd7fc-139">说明</span><span class="sxs-lookup"><span data-stu-id="fd7fc-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fd7fc-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="fd7fc-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="fd7fc-141">是</span><span class="sxs-lookup"><span data-stu-id="fd7fc-141">Yes</span></span>   |  <span data-ttu-id="fd7fc-142">定义桌面外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="fd7fc-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="fd7fc-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="fd7fc-144">否</span><span class="sxs-lookup"><span data-stu-id="fd7fc-144">No</span></span>   |  <span data-ttu-id="fd7fc-p103">定义移动外形规格的设置。**注意：** 仅在 Outlook for iOS 中支持此元素。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="fd7fc-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="fd7fc-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="fd7fc-148">No</span><span class="sxs-lookup"><span data-stu-id="fd7fc-148">No</span></span>   |  <span data-ttu-id="fd7fc-149">定义所有外形规格的设置。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="fd7fc-150">仅用于 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="fd7fc-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fd7fc-151">xsi:type</span></span>

<span data-ttu-id="fd7fc-152">控制所包含的设置适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。</span><span class="sxs-lookup"><span data-stu-id="fd7fc-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="fd7fc-153">值必须为以下值之一：</span><span class="sxs-lookup"><span data-stu-id="fd7fc-153">The value must be one of the following:</span></span>

- <span data-ttu-id="fd7fc-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-154">`Document` (Word)</span></span>
- <span data-ttu-id="fd7fc-155">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="fd7fc-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="fd7fc-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="fd7fc-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="fd7fc-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="fd7fc-159">主机示例</span><span class="sxs-lookup"><span data-stu-id="fd7fc-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
