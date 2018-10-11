
# <a name="context"></a><span data-ttu-id="ad941-101">context</span><span class="sxs-lookup"><span data-stu-id="ad941-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="ad941-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="ad941-102">[Office](Office.md).context</span></span>

<span data-ttu-id="ad941-p101">Office.context 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[共享 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="ad941-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad941-105">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-105">Requirements</span></span>

|<span data-ttu-id="ad941-106">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-106">Requirement</span></span>| <span data-ttu-id="ad941-107">值</span><span class="sxs-lookup"><span data-stu-id="ad941-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad941-108">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="ad941-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad941-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ad941-109">1.0</span></span>|
|[<span data-ttu-id="ad941-110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ad941-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ad941-111">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ad941-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ad941-112">成员和方法</span><span class="sxs-lookup"><span data-stu-id="ad941-112">Members and methods</span></span>

| <span data-ttu-id="ad941-113">成员</span><span class="sxs-lookup"><span data-stu-id="ad941-113">Member</span></span> | <span data-ttu-id="ad941-114">类型</span><span class="sxs-lookup"><span data-stu-id="ad941-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ad941-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ad941-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ad941-116">成员</span><span class="sxs-lookup"><span data-stu-id="ad941-116">Member</span></span> |
| [<span data-ttu-id="ad941-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="ad941-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="ad941-118">成员</span><span class="sxs-lookup"><span data-stu-id="ad941-118">Member</span></span> |
| [<span data-ttu-id="ad941-119">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ad941-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings) | <span data-ttu-id="ad941-120">成员</span><span class="sxs-lookup"><span data-stu-id="ad941-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ad941-121">命名空间</span><span class="sxs-lookup"><span data-stu-id="ad941-121">Namespaces</span></span>

<span data-ttu-id="ad941-122">[邮箱](office.context.mailbox.md)为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="ad941-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="ad941-123">成员</span><span class="sxs-lookup"><span data-stu-id="ad941-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="ad941-124">displayLanguage： 字符串</span><span class="sxs-lookup"><span data-stu-id="ad941-124">displayLanguage :String</span></span>

<span data-ttu-id="ad941-125">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="ad941-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="ad941-126">`displayLanguage` 值反映在 Office 主机应用程序中当前**显示语言**  设置通过 **文件 > 选项 > 语言**</span><span class="sxs-lookup"><span data-stu-id="ad941-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ad941-127">类型:</span><span class="sxs-lookup"><span data-stu-id="ad941-127">Type:</span></span>

*   <span data-ttu-id="ad941-128">字符串</span><span class="sxs-lookup"><span data-stu-id="ad941-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad941-129">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-129">Requirements</span></span>

|<span data-ttu-id="ad941-130">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-130">Requirement</span></span>| <span data-ttu-id="ad941-131">值</span><span class="sxs-lookup"><span data-stu-id="ad941-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad941-132">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="ad941-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad941-133">1.0</span><span class="sxs-lookup"><span data-stu-id="ad941-133">1.0</span></span>|
|[<span data-ttu-id="ad941-134">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ad941-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ad941-135">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ad941-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad941-136">示例</span><span class="sxs-lookup"><span data-stu-id="ad941-136">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="ad941-137">officeTheme： 对象</span><span class="sxs-lookup"><span data-stu-id="ad941-137">officeTheme :Object</span></span>

<span data-ttu-id="ad941-138">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="ad941-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="ad941-139">注意：在 iOS 版 Outlook 或  Android 版 Outlook 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="ad941-139">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ad941-p102">通过使用 Office 主题颜色，你可以使外接程序的配色方案与用户（通过 **“文件”>“Office 帐户”>“Office 主题”UI**）选择的当前 Office 主题协调一致，这种做法适用于所有 Office 主机应用程序。使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="ad941-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ad941-142">类型:</span><span class="sxs-lookup"><span data-stu-id="ad941-142">Type:</span></span>

*   <span data-ttu-id="ad941-143">类型</span><span class="sxs-lookup"><span data-stu-id="ad941-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="ad941-144">属性：</span><span class="sxs-lookup"><span data-stu-id="ad941-144">Properties:</span></span>

|<span data-ttu-id="ad941-145">名称</span><span class="sxs-lookup"><span data-stu-id="ad941-145">Name</span></span>| <span data-ttu-id="ad941-146">类型</span><span class="sxs-lookup"><span data-stu-id="ad941-146">Type</span></span>| <span data-ttu-id="ad941-147">说明</span><span class="sxs-lookup"><span data-stu-id="ad941-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="ad941-148">字符串</span><span class="sxs-lookup"><span data-stu-id="ad941-148">String</span></span>|<span data-ttu-id="ad941-149">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="ad941-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="ad941-150">字符串</span><span class="sxs-lookup"><span data-stu-id="ad941-150">String</span></span>|<span data-ttu-id="ad941-151">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="ad941-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="ad941-152">字符串</span><span class="sxs-lookup"><span data-stu-id="ad941-152">String</span></span>|<span data-ttu-id="ad941-153">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="ad941-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="ad941-154">字符串</span><span class="sxs-lookup"><span data-stu-id="ad941-154">String</span></span>|<span data-ttu-id="ad941-155">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="ad941-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad941-156">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-156">Requirements</span></span>

|<span data-ttu-id="ad941-157">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-157">Requirement</span></span>| <span data-ttu-id="ad941-158">值</span><span class="sxs-lookup"><span data-stu-id="ad941-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad941-159">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="ad941-159">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad941-160">1.3</span><span class="sxs-lookup"><span data-stu-id="ad941-160">1.3</span></span>|
|[<span data-ttu-id="ad941-161">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ad941-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ad941-162">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ad941-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ad941-163">示例</span><span class="sxs-lookup"><span data-stu-id="ad941-163">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="ad941-164">roamingSettings:[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="ad941-164">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="ad941-165">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="ad941-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ad941-166">对象`RoamingSettings` 允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="ad941-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ad941-167">类型:</span><span class="sxs-lookup"><span data-stu-id="ad941-167">Type:</span></span>

*   [<span data-ttu-id="ad941-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ad941-168">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ad941-169">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-169">Requirements</span></span>

|<span data-ttu-id="ad941-170">要求</span><span class="sxs-lookup"><span data-stu-id="ad941-170">Requirement</span></span>| <span data-ttu-id="ad941-171">值</span><span class="sxs-lookup"><span data-stu-id="ad941-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad941-172">最低的邮箱版本要求</span><span class="sxs-lookup"><span data-stu-id="ad941-172">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad941-173">1.0</span><span class="sxs-lookup"><span data-stu-id="ad941-173">1.0</span></span>|
|[<span data-ttu-id="ad941-174">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ad941-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ad941-175">受限</span><span class="sxs-lookup"><span data-stu-id="ad941-175">Restricted</span></span>|
|[<span data-ttu-id="ad941-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ad941-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ad941-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ad941-177">Compose or read</span></span>|