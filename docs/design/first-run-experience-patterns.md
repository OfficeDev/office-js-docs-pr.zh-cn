# <a name="first-run-experience-patterns"></a><span data-ttu-id="c30d7-101">首次运行体验模式</span><span class="sxs-lookup"><span data-stu-id="c30d7-101">First-run experience patterns</span></span>

<span data-ttu-id="c30d7-102">首次运行体验 (FRE) 是用户对您的外接程序的介绍。</span><span class="sxs-lookup"><span data-stu-id="c30d7-102">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="c30d7-103">当用户第一次打开外接程序时，系统将会显示 FRE，其中提供对相应外接程序的功能、特性和/或优势的见解。</span><span class="sxs-lookup"><span data-stu-id="c30d7-103">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="c30d7-104">该体验有助于形成用户对外接程序的印象，对他们返回并继续使用外接程序的可能性产生强烈影响。</span><span class="sxs-lookup"><span data-stu-id="c30d7-104">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="c30d7-105">最佳做法</span><span class="sxs-lookup"><span data-stu-id="c30d7-105">Best practices</span></span>


<span data-ttu-id="c30d7-106">在制定首次运行体验时，请遵循以下最佳做法：</span><span class="sxs-lookup"><span data-stu-id="c30d7-106">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="c30d7-107">允许事项</span><span class="sxs-lookup"><span data-stu-id="c30d7-107">Do</span></span>|<span data-ttu-id="c30d7-108">禁止事项</span><span class="sxs-lookup"><span data-stu-id="c30d7-108">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="c30d7-109">简明扼要地介绍外接程序中的主要操作。</span><span class="sxs-lookup"><span data-stu-id="c30d7-109">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="c30d7-110">不要包含与使用入门无关的信息和标注。</span><span class="sxs-lookup"><span data-stu-id="c30d7-110">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="c30d7-111">让用户有机会完成一项对他们使用外接程序产生积极影响的操作。</span><span class="sxs-lookup"><span data-stu-id="c30d7-111">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="c30d7-112">不要期望用户一次了解所有内容。</span><span class="sxs-lookup"><span data-stu-id="c30d7-112">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="c30d7-113">专注于提供最大价值的操作。</span><span class="sxs-lookup"><span data-stu-id="c30d7-113">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="c30d7-114">打造用户想要完成的有参与感的体验。</span><span class="sxs-lookup"><span data-stu-id="c30d7-114">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="c30d7-115">不要强迫用户点击首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="c30d7-115">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="c30d7-116">为用户提供绕过首次运行体验的选项。</span><span class="sxs-lookup"><span data-stu-id="c30d7-116">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="c30d7-117">考虑是一次还是定期多次向用户显示首次运行体验是否对你的方案非常重要。</span><span class="sxs-lookup"><span data-stu-id="c30d7-117">Consider whether showing users the first-run experience once or many times is important to your scenario.</span></span> <span data-ttu-id="c30d7-118">例如，如果用户仅定期使用您的外接程序，则用户可能会对外接程序不太熟悉，并且可能会受益于与首次运行体验的其他交互。</span><span class="sxs-lookup"><span data-stu-id="c30d7-118">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="c30d7-119">根据适用情况应用以下模式，以创建或增强外接程序的首次运行体验。</span><span class="sxs-lookup"><span data-stu-id="c30d7-119">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="c30d7-120">旋转木马</span><span class="sxs-lookup"><span data-stu-id="c30d7-120">Carousel</span></span>


<span data-ttu-id="c30d7-121">在开始使用外接程序之前，旋转木马向用户展示一系列功能或信息。</span><span class="sxs-lookup"><span data-stu-id="c30d7-121">Walkthrough takes users through a series of features or information before they start using the add-in. (PDF, code)</span></span>

<span data-ttu-id="c30d7-122">*图 1：允许用户前进或跳过旋转木马流程的开始页面。*
![首次运行 - 旋转木马 - 桌面任务窗格规范](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="c30d7-122">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="c30d7-123">*图 2：仅根据有效传递信息的需要，最大限度减少您向用户展示的旋转木马屏幕数量*
![首次运行 - 旋转木马 - 桌面任务窗格规范](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="c30d7-123">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="c30d7-124">*图 3：提供明确号召性用语，以退出首次运行体验。*
![首次运行 - 旋转木马 - 桌面任务窗格规范](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="c30d7-124">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="c30d7-125">价值餐具垫</span><span class="sxs-lookup"><span data-stu-id="c30d7-125">Value Placemat</span></span>

<span data-ttu-id="c30d7-126">价值展示通过徽标展示位置、明确阐明的价值主张、功能亮点或摘要以及号召性用语来传达您的外接程序的价值主张。</span><span class="sxs-lookup"><span data-stu-id="c30d7-126">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="c30d7-127">![首次运行 - 价值餐具垫 - 桌面任务窗格规范](../images/add-in-FRE-value.png)
*带有徽标的价值餐具垫、明确的价值主张，功能摘要和号召性用语。*</span><span class="sxs-lookup"><span data-stu-id="c30d7-127">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="c30d7-128">视频餐具垫</span><span class="sxs-lookup"><span data-stu-id="c30d7-128">Video Placemat</span></span>

<span data-ttu-id="c30d7-129">视频餐具垫在用户开始使用你的外接程序之前向其展示视频。</span><span class="sxs-lookup"><span data-stu-id="c30d7-129">Video shows users a video before they start using your add-in. (spec, code)</span></span>


<span data-ttu-id="c30d7-130">*图 1：首次运行餐具垫 - 屏幕包含带有播放按钮和召性用语按钮的视频静止图像。*![视频餐具垫 - 桌面任务窗格规范](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="c30d7-130">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="c30d7-131">*图 2：视频播放器 - 用户会在对话框窗口中看到一段视频。*
![视频餐具垫 - 桌面任务窗格规范](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="c30d7-131">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
