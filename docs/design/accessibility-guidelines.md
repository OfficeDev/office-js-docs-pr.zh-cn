# <a name="accessibility-guidelines"></a><span data-ttu-id="8b55a-101">辅助功能准则</span><span class="sxs-lookup"><span data-stu-id="8b55a-101">Accessibility guidelines</span></span>

<span data-ttu-id="8b55a-p101">在设计和开发 Office 外接程序时，你将需要确保所有潜在用户和客户都能够成功使用你的外接程序。请应用以下准则，确保所有目标用户都能访问你的解决方案。</span><span class="sxs-lookup"><span data-stu-id="8b55a-p101">As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.</span></span>

## <a name="design-for-multiple-input-methods"></a><span data-ttu-id="8b55a-104">针对多种输入方式的设计</span><span class="sxs-lookup"><span data-stu-id="8b55a-104">Design for multiple input methods</span></span>

- <span data-ttu-id="8b55a-p102">确保用户可以仅通过键盘执行操作。用户应该能够使用 Tab 键和箭头键组合移动到页面上的所有可操作元素。</span><span class="sxs-lookup"><span data-stu-id="8b55a-p102">Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.</span></span>
- <span data-ttu-id="8b55a-107">在移动设备上，当用户通过触摸操作某个控件时，设备应该提供有用的音频反馈。</span><span class="sxs-lookup"><span data-stu-id="8b55a-107">On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.</span></span>
- <span data-ttu-id="8b55a-108">为所有交互式控件提供有用的标签。</span><span class="sxs-lookup"><span data-stu-id="8b55a-108">Provide helpful labels for all interactive controls.</span></span> 

## <a name="make-your-add-in-easy-to-use"></a><span data-ttu-id="8b55a-109">使你的外接程序易于使用</span><span class="sxs-lookup"><span data-stu-id="8b55a-109">Make your add-in easy to use</span></span>

- <span data-ttu-id="8b55a-110">不依赖于单个属性（例如颜色、大小、形状、位置、方向或声音）在 UI 中传达含义。</span><span class="sxs-lookup"><span data-stu-id="8b55a-110">Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.</span></span>
- <span data-ttu-id="8b55a-111">避免对上下文的意外更改，例如在用户未操作的情况下将焦点移到其他 UI 元素。</span><span class="sxs-lookup"><span data-stu-id="8b55a-111">Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.</span></span>
- <span data-ttu-id="8b55a-112">提供验证、确认或撤消所有绑定操作的方法。</span><span class="sxs-lookup"><span data-stu-id="8b55a-112">Provide a way to verify, confirm, or reverse all binding actions.</span></span>
- <span data-ttu-id="8b55a-113">提供暂停或停止媒体（例如音频和视频）的方法。</span><span class="sxs-lookup"><span data-stu-id="8b55a-113">Provide a way to pause or stop media, such as audio and video.</span></span>
- <span data-ttu-id="8b55a-114">不对用户操作施加时间限制。</span><span class="sxs-lookup"><span data-stu-id="8b55a-114">Do not impose a time limit for user action.</span></span>

## <a name="make-your-add-in-easy-to-see"></a><span data-ttu-id="8b55a-115">使你的外接程序易于查看</span><span class="sxs-lookup"><span data-stu-id="8b55a-115">Make your add-in easy to see</span></span>

- <span data-ttu-id="8b55a-116">避免对颜色的意外更改。</span><span class="sxs-lookup"><span data-stu-id="8b55a-116">Avoid unexpected color changes.</span></span>
- <span data-ttu-id="8b55a-p103">提供有意义且及时的信息，以描述 UI 元素、标题和标头、输入和错误。确保控件名称恰当地描述了控件的意图。</span><span class="sxs-lookup"><span data-stu-id="8b55a-p103">Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.</span></span>
- <span data-ttu-id="8b55a-119">遵循针对颜色对比度的[标准准则](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html)。</span><span class="sxs-lookup"><span data-stu-id="8b55a-119">Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.</span></span>

## <a name="account-for-assistive-technologies"></a><span data-ttu-id="8b55a-120">对辅助技术负责</span><span class="sxs-lookup"><span data-stu-id="8b55a-120">Account for assistive technologies</span></span>

- <span data-ttu-id="8b55a-121">避免使用会干扰辅助技术的功能，例如视觉、音频或其他交互。</span><span class="sxs-lookup"><span data-stu-id="8b55a-121">Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.</span></span>
- <span data-ttu-id="8b55a-p104">请勿以图像格式提供文本。屏幕阅读器无法读取图像中的文本。</span><span class="sxs-lookup"><span data-stu-id="8b55a-p104">Do not provide text in an image format. Screen readers cannot read text within images.</span></span>
- <span data-ttu-id="8b55a-124">为用户提供调整或静音所有音频源的方法。</span><span class="sxs-lookup"><span data-stu-id="8b55a-124">Provide a way for users to adjust or mute all audio sources.</span></span>
- <span data-ttu-id="8b55a-125">为用户提供打开字幕或音频说明与音频源的方法。</span><span class="sxs-lookup"><span data-stu-id="8b55a-125">Provide a way for users to turn on captions or audio description with audio sources.</span></span>
- <span data-ttu-id="8b55a-126">提供用于预警用户的声音替代方法，如视觉提示或振动。</span><span class="sxs-lookup"><span data-stu-id="8b55a-126">Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.</span></span>

## <a name="see-also"></a><span data-ttu-id="8b55a-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8b55a-127">See also</span></span>

- [<span data-ttu-id="8b55a-128">Web 内容辅助功能指南 (WCAG) 2.0</span><span class="sxs-lookup"><span data-stu-id="8b55a-128">Web Content Accessibility Guidelines (WCAG) 2.0</span></span>](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [<span data-ttu-id="8b55a-129">向非 Web 信息和通信技术 (WCAG2ICT) 应用 WCAG 2.0 的指南</span><span class="sxs-lookup"><span data-stu-id="8b55a-129">Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)</span></span>](https://www.w3.org/TR/wcag2ict/)
- [<span data-ttu-id="8b55a-130">关于信息和通信技术 (ICT) 的辅助功能要求的欧洲标准</span><span class="sxs-lookup"><span data-stu-id="8b55a-130">European Standard on accessibility requirements for Information and Communication Technologies (ICT)</span></span>](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
