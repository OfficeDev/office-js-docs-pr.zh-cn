# <a name="requestedheight-element"></a><span data-ttu-id="c8156-101">RequestedHeight 元素</span><span class="sxs-lookup"><span data-stu-id="c8156-101">RequestedHeight element</span></span>

<span data-ttu-id="c8156-102">指定内容加载项或邮件加载项的初始高度（以像素为单位）。</span><span class="sxs-lookup"><span data-stu-id="c8156-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="c8156-103">**加载项类型：** 内容、邮件</span><span class="sxs-lookup"><span data-stu-id="c8156-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c8156-104">语法</span><span class="sxs-lookup"><span data-stu-id="c8156-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="c8156-105">包含在</span><span class="sxs-lookup"><span data-stu-id="c8156-105">Contained in:</span></span>

- <span data-ttu-id="c8156-106">[DefaultSettings](defaultsettings.md)（内容加载项）可以是介于 32 和 1000 之间的值</span><span class="sxs-lookup"><span data-stu-id="c8156-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="c8156-107">[DesktopSettings](desktopsettings.md) 和 [TabletSettings](tabletsettings.md)（邮件加载项）可以是介于 32 和 450 之间的值</span><span class="sxs-lookup"><span data-stu-id="c8156-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="c8156-108">[ExtensionPoint](extensionpoint.md)（上下文邮件加载项）可以是介于 140 和 450（对于 **DetectedEntity** 扩展点）及介于 32 和 450（**CustomPane** 扩展点）之间的值</span><span class="sxs-lookup"><span data-stu-id="c8156-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>