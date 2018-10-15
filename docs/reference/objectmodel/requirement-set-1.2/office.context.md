
# <a name="context"></a><span data-ttu-id="cfd94-101">context</span><span class="sxs-lookup"><span data-stu-id="cfd94-101">context</span></span>

### <span data-ttu-id="cfd94-p101">[Office](Office.md). context</span><span class="sxs-lookup"><span data-stu-id="cfd94-p101">[Office](Office.md). context</span></span>

<span data-ttu-id="cfd94-p102">Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[共享 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="cfd94-p102">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="cfd94-106">要求</span><span class="sxs-lookup"><span data-stu-id="cfd94-106">Requirements</span></span>

|<span data-ttu-id="cfd94-107">要求</span><span class="sxs-lookup"><span data-stu-id="cfd94-107">Requirement</span></span>| <span data-ttu-id="cfd94-108">值</span><span class="sxs-lookup"><span data-stu-id="cfd94-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cfd94-109">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="cfd94-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cfd94-110">1.0</span><span class="sxs-lookup"><span data-stu-id="cfd94-110">1.0</span></span>|
|[<span data-ttu-id="cfd94-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cfd94-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cfd94-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cfd94-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="cfd94-113">命名空间</span><span class="sxs-lookup"><span data-stu-id="cfd94-113">Namespaces</span></span>

<span data-ttu-id="cfd94-114">[mailbox](office.context.mailbox.md)：为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 加载项对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="cfd94-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="cfd94-115">成员</span><span class="sxs-lookup"><span data-stu-id="cfd94-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="cfd94-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="cfd94-116">displayLanguage :String</span></span>

<span data-ttu-id="cfd94-117">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="cfd94-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="cfd94-118">`displayLanguage` 值反映在 Office 主机应用程序中通过**文件 > 选项 > 语言**设置的当前**显示语言**。</span><span class="sxs-lookup"><span data-stu-id="cfd94-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="cfd94-119">类型：</span><span class="sxs-lookup"><span data-stu-id="cfd94-119">Type:</span></span>

*   <span data-ttu-id="cfd94-120">字符串</span><span class="sxs-lookup"><span data-stu-id="cfd94-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cfd94-121">要求</span><span class="sxs-lookup"><span data-stu-id="cfd94-121">Requirements</span></span>

|<span data-ttu-id="cfd94-122">要求</span><span class="sxs-lookup"><span data-stu-id="cfd94-122">Requirement</span></span>| <span data-ttu-id="cfd94-123">值</span><span class="sxs-lookup"><span data-stu-id="cfd94-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="cfd94-124">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="cfd94-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cfd94-125">1.0</span><span class="sxs-lookup"><span data-stu-id="cfd94-125">1.0</span></span>|
|[<span data-ttu-id="cfd94-126">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cfd94-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cfd94-127">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cfd94-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="cfd94-128">示例</span><span class="sxs-lookup"><span data-stu-id="cfd94-128">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a><span data-ttu-id="cfd94-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="cfd94-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span></span>

<span data-ttu-id="cfd94-130">获取一个对象，它表示保存到用户邮箱的邮件加载项的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="cfd94-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="cfd94-131">对象 `RoamingSettings` 允许您存储和访问用户邮箱中存储的邮件加载项的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该加载项时，加载项可以使用数据。</span><span class="sxs-lookup"><span data-stu-id="cfd94-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="cfd94-132">类型：</span><span class="sxs-lookup"><span data-stu-id="cfd94-132">Type:</span></span>

*   [<span data-ttu-id="cfd94-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cfd94-133">RoamingSettings</span></span>](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="cfd94-134">要求</span><span class="sxs-lookup"><span data-stu-id="cfd94-134">Requirements</span></span>

|<span data-ttu-id="cfd94-135">要求</span><span class="sxs-lookup"><span data-stu-id="cfd94-135">Requirement</span></span>| <span data-ttu-id="cfd94-136">值</span><span class="sxs-lookup"><span data-stu-id="cfd94-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="cfd94-137">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="cfd94-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cfd94-138">1.0</span><span class="sxs-lookup"><span data-stu-id="cfd94-138">1.0</span></span>|
|[<span data-ttu-id="cfd94-139">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cfd94-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cfd94-140">受限</span><span class="sxs-lookup"><span data-stu-id="cfd94-140">Restricted</span></span>|
|[<span data-ttu-id="cfd94-141">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cfd94-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cfd94-142">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cfd94-142">Compose or read</span></span>|