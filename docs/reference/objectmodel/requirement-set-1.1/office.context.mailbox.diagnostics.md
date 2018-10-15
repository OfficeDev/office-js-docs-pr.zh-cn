
# <a name="diagnostics"></a><span data-ttu-id="a72a9-101">诊断</span><span class="sxs-lookup"><span data-stu-id="a72a9-101">diagnostics</span></span>

### <span data-ttu-id="a72a9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="a72a9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="a72a9-104">将诊断信息提供给 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="a72a9-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a72a9-105">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-105">Requirements</span></span>

|<span data-ttu-id="a72a9-106">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-106">Requirement</span></span>| <span data-ttu-id="a72a9-107">值</span><span class="sxs-lookup"><span data-stu-id="a72a9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a72a9-108">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a72a9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a72a9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a72a9-109">1.0</span></span>|
|[<span data-ttu-id="a72a9-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a72a9-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a72a9-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a72a9-111">ReadItem</span></span>|
|[<span data-ttu-id="a72a9-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a72a9-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a72a9-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a72a9-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="a72a9-114">成员</span><span class="sxs-lookup"><span data-stu-id="a72a9-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="a72a9-115">hostName :字符串</span><span class="sxs-lookup"><span data-stu-id="a72a9-115">hostName :String</span></span>

<span data-ttu-id="a72a9-116">获取表示主机应用程序的名称的字符串。</span><span class="sxs-lookup"><span data-stu-id="a72a9-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="a72a9-117">字符串可以是下列值之一：`Outlook`、`OutlookIOS` 或 `OutlookWebApp`。</span><span class="sxs-lookup"><span data-stu-id="a72a9-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="a72a9-118">类型：</span><span class="sxs-lookup"><span data-stu-id="a72a9-118">Type:</span></span>

*   <span data-ttu-id="a72a9-119">String</span><span class="sxs-lookup"><span data-stu-id="a72a9-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a72a9-120">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-120">Requirements</span></span>

|<span data-ttu-id="a72a9-121">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-121">Requirement</span></span>| <span data-ttu-id="a72a9-122">值</span><span class="sxs-lookup"><span data-stu-id="a72a9-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="a72a9-123">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a72a9-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a72a9-124">1.0</span><span class="sxs-lookup"><span data-stu-id="a72a9-124">1.0</span></span>|
|[<span data-ttu-id="a72a9-125">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a72a9-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a72a9-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a72a9-126">ReadItem</span></span>|
|[<span data-ttu-id="a72a9-127">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a72a9-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a72a9-128">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a72a9-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="a72a9-129">hostVersion :字符串</span><span class="sxs-lookup"><span data-stu-id="a72a9-129">hostVersion :String</span></span>

<span data-ttu-id="a72a9-130">获取表示主机应用程序或 Exchange Server 的版本的字符串。</span><span class="sxs-lookup"><span data-stu-id="a72a9-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="a72a9-p102">如果邮件加载项是在 Outlook 桌面客户端或 Outlook for iOS 上运行，则 `hostVersion` 属性返回主机应用程序 Outlook 的版本。在 Outlook Web App 中，属性返回 Exchange Server 的版本。其中的一个示例是字符串 `15.0.468.0`。</span><span class="sxs-lookup"><span data-stu-id="a72a9-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="a72a9-134">类型：</span><span class="sxs-lookup"><span data-stu-id="a72a9-134">Type:</span></span>

*   <span data-ttu-id="a72a9-135">String</span><span class="sxs-lookup"><span data-stu-id="a72a9-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a72a9-136">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-136">Requirements</span></span>

|<span data-ttu-id="a72a9-137">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-137">Requirement</span></span>| <span data-ttu-id="a72a9-138">值</span><span class="sxs-lookup"><span data-stu-id="a72a9-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="a72a9-139">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a72a9-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a72a9-140">1.0</span><span class="sxs-lookup"><span data-stu-id="a72a9-140">1.0</span></span>|
|[<span data-ttu-id="a72a9-141">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a72a9-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a72a9-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a72a9-142">ReadItem</span></span>|
|[<span data-ttu-id="a72a9-143">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a72a9-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a72a9-144">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a72a9-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="a72a9-145">OWAView :字符串</span><span class="sxs-lookup"><span data-stu-id="a72a9-145">OWAView :String</span></span>

<span data-ttu-id="a72a9-146">获取表示 Outlook Web App 的当前视图的字符串。</span><span class="sxs-lookup"><span data-stu-id="a72a9-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="a72a9-147">返回的字符串可以是下列值之一：`OneColumn`、`TwoColumns` 或 `ThreeColumns`。</span><span class="sxs-lookup"><span data-stu-id="a72a9-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="a72a9-148">如果主机应用程序不是 Outlook Web App，则访问此属性将导致返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="a72a9-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="a72a9-149">Outlook Web App 具有三种视图，这些视图分别与屏幕和窗口的宽度以及可以显示的列数相对应：</span><span class="sxs-lookup"><span data-stu-id="a72a9-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="a72a9-p103">`OneColumn`在屏幕为窄屏时显示。Outlook Web App 在智能手机的整个屏幕上使用此单列布局。</span><span class="sxs-lookup"><span data-stu-id="a72a9-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="a72a9-p104">`TwoColumns`在屏幕较宽时显示。Outlook Web App 在大多数平板电脑上使用此视图。</span><span class="sxs-lookup"><span data-stu-id="a72a9-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="a72a9-p105">`ThreeColumns`在屏幕为宽屏时显示。例如，Outlook Web App 在台式计算机的全屏窗口中使用此视图。</span><span class="sxs-lookup"><span data-stu-id="a72a9-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="a72a9-156">类型：</span><span class="sxs-lookup"><span data-stu-id="a72a9-156">Type:</span></span>

*   <span data-ttu-id="a72a9-157">String</span><span class="sxs-lookup"><span data-stu-id="a72a9-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a72a9-158">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-158">Requirements</span></span>

|<span data-ttu-id="a72a9-159">要求</span><span class="sxs-lookup"><span data-stu-id="a72a9-159">Requirement</span></span>| <span data-ttu-id="a72a9-160">值</span><span class="sxs-lookup"><span data-stu-id="a72a9-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a72a9-161">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="a72a9-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a72a9-162">1.0</span><span class="sxs-lookup"><span data-stu-id="a72a9-162">1.0</span></span>|
|[<span data-ttu-id="a72a9-163">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="a72a9-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a72a9-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a72a9-164">ReadItem</span></span>|
|[<span data-ttu-id="a72a9-165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="a72a9-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a72a9-166">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="a72a9-166">Compose or read</span></span>|