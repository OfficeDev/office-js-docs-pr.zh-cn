# <a name="getstarted-element"></a><span data-ttu-id="6badd-101">GetStarted 元素</span><span class="sxs-lookup"><span data-stu-id="6badd-101">GetStarted element</span></span>

<span data-ttu-id="6badd-p101">提供在 Word、Excel、PowerPoint 和 OneNote 主机中安装加载项时显示的标注所使用的信息。**GetStarted** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="6badd-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="6badd-104">子元素</span><span class="sxs-lookup"><span data-stu-id="6badd-104">Child elements</span></span>

| <span data-ttu-id="6badd-105">元素</span><span class="sxs-lookup"><span data-stu-id="6badd-105">Element</span></span>                       | <span data-ttu-id="6badd-106">必需</span><span class="sxs-lookup"><span data-stu-id="6badd-106">Required</span></span> | <span data-ttu-id="6badd-107">说明</span><span class="sxs-lookup"><span data-stu-id="6badd-107">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="6badd-108">Title</span><span class="sxs-lookup"><span data-stu-id="6badd-108">Title</span></span>](#title)               | <span data-ttu-id="6badd-109">是</span><span class="sxs-lookup"><span data-stu-id="6badd-109">Yes</span></span>      | <span data-ttu-id="6badd-110">定义加载项公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="6badd-110">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="6badd-111">说明</span><span class="sxs-lookup"><span data-stu-id="6badd-111">Description</span></span>](#description)   | <span data-ttu-id="6badd-112">是</span><span class="sxs-lookup"><span data-stu-id="6badd-112">Yes</span></span>      | <span data-ttu-id="6badd-113">包含 JavaScript 函数的文件的 URL。</span><span class="sxs-lookup"><span data-stu-id="6badd-113">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="6badd-114">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="6badd-114">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="6badd-115">否</span><span class="sxs-lookup"><span data-stu-id="6badd-115">No</span></span>       | <span data-ttu-id="6badd-116">指向详细说明加载项的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="6badd-116">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="6badd-117">Title</span><span class="sxs-lookup"><span data-stu-id="6badd-117">Title</span></span> 

<span data-ttu-id="6badd-p102">必需。用于标注顶部的标题。**resid** 属性引用 [Resources](resources.md) 分区的 **ShortStrings** 元素中的有效 ID。</span><span class="sxs-lookup"><span data-stu-id="6badd-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="6badd-121">说明</span><span class="sxs-lookup"><span data-stu-id="6badd-121">Description</span></span>

<span data-ttu-id="6badd-p103">必需。标注的说明/正文内容。**resid** 属性引用 [Resources](resources.md) 分区的 **LongStrings** 元素中的有效 ID。</span><span class="sxs-lookup"><span data-stu-id="6badd-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="6badd-125">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="6badd-125">LearnMoreUrl</span></span>

<span data-ttu-id="6badd-p104">必需。指向用户可以了解你的外接程序详细信息的页面 URL。**resid** 属性引用 [Resources](resources.md) 分区的 **Urls** 元素中的有效 ID。</span><span class="sxs-lookup"><span data-stu-id="6badd-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="6badd-129">**LearnMoreUrl** 当前无法在 Word、Excel 或 PowerPoint 客户端中呈现。</span><span class="sxs-lookup"><span data-stu-id="6badd-129">NOTE:**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="6badd-130">我们建议为所有客户端添加此 URL，以便 URL 在可用时呈现。</span><span class="sxs-lookup"><span data-stu-id="6badd-130">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="6badd-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6badd-131">See also</span></span>

<span data-ttu-id="6badd-132">下面的代码示例使用 **GetStarted** 元素：</span><span class="sxs-lookup"><span data-stu-id="6badd-132">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="6badd-133">用于控制表和图表格式化的 Excel Web 加载项</span><span class="sxs-lookup"><span data-stu-id="6badd-133">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="6badd-134">Word 加载项 JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="6badd-134">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="6badd-135">在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表</span><span class="sxs-lookup"><span data-stu-id="6badd-135">Insert Excel charts using Microsoft Graph in a PowerPoint Add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
