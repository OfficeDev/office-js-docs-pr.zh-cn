# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="b96e6-101">适用于 Office 外接程序的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="b96e6-101">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="b96e6-102">Office加载项用户体验设计应该向Office用户提供引人注目的体验，并在默认的Office用户界面中，实现无缝配合，扩展Office整体体验。</span><span class="sxs-lookup"><span data-stu-id="b96e6-102">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="b96e6-103">我们的UX模式由各个组件组成。</span><span class="sxs-lookup"><span data-stu-id="b96e6-103">Our UX patterns are composed of components.</span></span> <span data-ttu-id="b96e6-104">组件是帮助客户与软件或服务元素进行交互的控件。</span><span class="sxs-lookup"><span data-stu-id="b96e6-104">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="b96e6-105">按钮、导航和菜单是常见组件的示例，通常具有一致的样式和行为。</span><span class="sxs-lookup"><span data-stu-id="b96e6-105">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="b96e6-106">Office UI Fabric 呈现外观和行为类似于 Office 部件的组件。</span><span class="sxs-lookup"><span data-stu-id="b96e6-106">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="b96e6-107">利用Fabric，易于与Office集成。</span><span class="sxs-lookup"><span data-stu-id="b96e6-107">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="b96e6-108">如果加载项自身预先存在组件语言，则不需要为了Fabric而放弃此语言。</span><span class="sxs-lookup"><span data-stu-id="b96e6-108">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="b96e6-109">与 Office 集成的同时寻找保留该语言的机会。</span><span class="sxs-lookup"><span data-stu-id="b96e6-109">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="b96e6-110">寻找置换出风格元素、删除冲突，或采用样式和行为以避免用户混淆的方法。</span><span class="sxs-lookup"><span data-stu-id="b96e6-110">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="b96e6-111">所提供的模式是基于常见客户场景和用户体验研究的最佳实践解决方案。</span><span class="sxs-lookup"><span data-stu-id="b96e6-111">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="b96e6-112">它们旨在为设计和开发加载项提供快速切入点，并为实现微软和品牌元素之间的平衡提供指导。</span><span class="sxs-lookup"><span data-stu-id="b96e6-112">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="b96e6-113">提供明快、现代化的用户体验，在此体验中，来自微软Fabric设计语言的设计元素与合作伙伴独特的品牌标识处于平衡状态，这可能有助于促使用户保留和采用您的加载项。</span><span class="sxs-lookup"><span data-stu-id="b96e6-113">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="b96e6-114">使用UX模式模板：</span><span class="sxs-lookup"><span data-stu-id="b96e6-114">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="b96e6-115">将解决方案应用于常见的客户方案。</span><span class="sxs-lookup"><span data-stu-id="b96e6-115">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="b96e6-116">应用设计最佳实践。</span><span class="sxs-lookup"><span data-stu-id="b96e6-116">Apply design best practices.</span></span>
* <span data-ttu-id="b96e6-117">纳入“[Office UI Fabric](https://developer.microsoft.com/fabric#/get-started)”组件和样式。</span><span class="sxs-lookup"><span data-stu-id="b96e6-117">Incorporate [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="b96e6-118">构建以可视方式与默认 Office UI 集成的外接程序。</span><span class="sxs-lookup"><span data-stu-id="b96e6-118">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="b96e6-119">想象UX。</span><span class="sxs-lookup"><span data-stu-id="b96e6-119">Ideate and visualize UX.</span></span>


## <a name="getting-started"></a><span data-ttu-id="b96e6-120">入门</span><span class="sxs-lookup"><span data-stu-id="b96e6-120">Getting started</span></span>

<span data-ttu-id="b96e6-121">这些模式按照加载项中常见的关键操作或体验进行组织。</span><span class="sxs-lookup"><span data-stu-id="b96e6-121">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="b96e6-122">主要组别是：</span><span class="sxs-lookup"><span data-stu-id="b96e6-122">The main groups are:</span></span>

* [<span data-ttu-id="b96e6-123">首次运行体验（FRE）</span><span class="sxs-lookup"><span data-stu-id="b96e6-123">First run experience</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="b96e6-124">身份验证</span><span class="sxs-lookup"><span data-stu-id="b96e6-124">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="b96e6-125">导航</span><span class="sxs-lookup"><span data-stu-id="b96e6-125">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="b96e6-126">品牌设计</span><span class="sxs-lookup"><span data-stu-id="b96e6-126">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="b96e6-127">浏览每个分组，了解如何使用最佳做法设计加载项。</span><span class="sxs-lookup"><span data-stu-id="b96e6-127">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>



><span data-ttu-id="b96e6-128">注：本文档所显示的示例屏幕按照 **1366×768**的分辨率进行设计和显示。</span><span class="sxs-lookup"><span data-stu-id="b96e6-128">NOTE: The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**</span></span>




## <a name="see-also"></a><span data-ttu-id="b96e6-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b96e6-129">See also</span></span>
* [<span data-ttu-id="b96e6-130">设计工具包</span><span class="sxs-lookup"><span data-stu-id="b96e6-130">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="b96e6-131">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="b96e6-131">Office UI Fabric</span></span>](https://developer.microsoft.com/fabric)
* [<span data-ttu-id="b96e6-132">Office 外接程序开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="b96e6-132">Best practices for developing Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-best-practices)
* [<span data-ttu-id="b96e6-133">开始使用 Fabric React</span><span class="sxs-lookup"><span data-stu-id="b96e6-133">name: Get started using Fabric React href: design/using-office-ui-fabric-react.md</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/using-office-ui-fabric-react)
