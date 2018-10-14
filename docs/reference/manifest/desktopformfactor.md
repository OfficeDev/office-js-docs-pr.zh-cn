# <a name="desktopformfactor-element"></a><span data-ttu-id="a928f-101">DesktopFormFactor 元素</span><span class="sxs-lookup"><span data-stu-id="a928f-101">DesktopFormFactor element</span></span>

<span data-ttu-id="a928f-p101">指定对桌面外形规格的加载项设置。桌面外形规格包括 Office for Windows、Office for Mac 和 Office Online。它包含桌面外形规格的所有加载项信息（**Resource** 节点除外）。</span><span class="sxs-lookup"><span data-stu-id="a928f-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="a928f-p102">每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。</span><span class="sxs-lookup"><span data-stu-id="a928f-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a928f-107">SupportsSharedFolders 元素仅在针对 Exchange Online 的 Outlook 加载项预览要求集中可用。</span><span class="sxs-lookup"><span data-stu-id="a928f-107">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span>
> <span data-ttu-id="a928f-108">使用此元素的加载项无法出现在 Office 商店或集中部署中。</span><span class="sxs-lookup"><span data-stu-id="a928f-108">Add-ins that use this element aren't allowed in the Office Store or Centralized Deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="a928f-109">子元素</span><span class="sxs-lookup"><span data-stu-id="a928f-109">Child elements</span></span>

| <span data-ttu-id="a928f-110">元素</span><span class="sxs-lookup"><span data-stu-id="a928f-110">Element</span></span>                               | <span data-ttu-id="a928f-111">必需</span><span class="sxs-lookup"><span data-stu-id="a928f-111">Required</span></span> | <span data-ttu-id="a928f-112">说明</span><span class="sxs-lookup"><span data-stu-id="a928f-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="a928f-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="a928f-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="a928f-114">是</span><span class="sxs-lookup"><span data-stu-id="a928f-114">Yes</span></span>      | <span data-ttu-id="a928f-115">定义加载项公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="a928f-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="a928f-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="a928f-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="a928f-117">是</span><span class="sxs-lookup"><span data-stu-id="a928f-117">Yes</span></span>      | <span data-ttu-id="a928f-118">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="a928f-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="a928f-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="a928f-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="a928f-120">否</span><span class="sxs-lookup"><span data-stu-id="a928f-120">No</span></span>       | <span data-ttu-id="a928f-121">定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。</span><span class="sxs-lookup"><span data-stu-id="a928f-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| <span data-ttu-id="a928f-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="a928f-122">SupportsSharedFolders</span></span>                 | <span data-ttu-id="a928f-123">否</span><span class="sxs-lookup"><span data-stu-id="a928f-123">No</span></span>       | <span data-ttu-id="a928f-124">定义 Outlook 加载项在委派方案中是否可用，默认设置为 *false* 。</span><span class="sxs-lookup"><span data-stu-id="a928f-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> <span data-ttu-id="a928f-125">预览要求集</span><span class="sxs-lookup"><span data-stu-id="a928f-125">Outlook add-in API Preview requirement set</span></span>|

## <a name="desktopformfactor-example"></a><span data-ttu-id="a928f-126">DesktopFormFactor 示例</span><span class="sxs-lookup"><span data-stu-id="a928f-126">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
