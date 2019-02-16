---
title: Office.context - 要求集 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 2d17d1a91d4ecbf8a46d0aa9c972187b0006fb82
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068313"
---
# <a name="context"></a><span data-ttu-id="5cbec-102">context</span><span class="sxs-lookup"><span data-stu-id="5cbec-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="5cbec-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="5cbec-103">[Office](Office.md).context</span></span>

<span data-ttu-id="5cbec-p101">Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[通用 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="5cbec-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5cbec-106">要求</span><span class="sxs-lookup"><span data-stu-id="5cbec-106">Requirements</span></span>

|<span data-ttu-id="5cbec-107">要求</span><span class="sxs-lookup"><span data-stu-id="5cbec-107">Requirement</span></span>| <span data-ttu-id="5cbec-108">值</span><span class="sxs-lookup"><span data-stu-id="5cbec-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cbec-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5cbec-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5cbec-110">1.0</span><span class="sxs-lookup"><span data-stu-id="5cbec-110">1.0</span></span>|
|[<span data-ttu-id="5cbec-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5cbec-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5cbec-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5cbec-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5cbec-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="5cbec-113">Members and methods</span></span>

| <span data-ttu-id="5cbec-114">成员</span><span class="sxs-lookup"><span data-stu-id="5cbec-114">Member</span></span> | <span data-ttu-id="5cbec-115">类型</span><span class="sxs-lookup"><span data-stu-id="5cbec-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5cbec-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="5cbec-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="5cbec-117">成员</span><span class="sxs-lookup"><span data-stu-id="5cbec-117">Member</span></span> |
| [<span data-ttu-id="5cbec-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="5cbec-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="5cbec-119">成员</span><span class="sxs-lookup"><span data-stu-id="5cbec-119">Member</span></span> |
| [<span data-ttu-id="5cbec-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="5cbec-120">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings) | <span data-ttu-id="5cbec-121">成员</span><span class="sxs-lookup"><span data-stu-id="5cbec-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5cbec-122">命名空间</span><span class="sxs-lookup"><span data-stu-id="5cbec-122">Namespaces</span></span>

<span data-ttu-id="5cbec-123">[mailbox](office.context.mailbox.md)：为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="5cbec-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="5cbec-124">成员</span><span class="sxs-lookup"><span data-stu-id="5cbec-124">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="5cbec-125">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="5cbec-125">displayLanguage :String</span></span>

<span data-ttu-id="5cbec-126">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="5cbec-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="5cbec-127">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="5cbec-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="5cbec-128">Type</span><span class="sxs-lookup"><span data-stu-id="5cbec-128">Type</span></span>

*   <span data-ttu-id="5cbec-129">String</span><span class="sxs-lookup"><span data-stu-id="5cbec-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5cbec-130">Requirements</span><span class="sxs-lookup"><span data-stu-id="5cbec-130">Requirements</span></span>

|<span data-ttu-id="5cbec-131">要求</span><span class="sxs-lookup"><span data-stu-id="5cbec-131">Requirement</span></span>| <span data-ttu-id="5cbec-132">值</span><span class="sxs-lookup"><span data-stu-id="5cbec-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cbec-133">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5cbec-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5cbec-134">1.0</span><span class="sxs-lookup"><span data-stu-id="5cbec-134">1.0</span></span>|
|[<span data-ttu-id="5cbec-135">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5cbec-135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5cbec-136">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5cbec-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5cbec-137">示例</span><span class="sxs-lookup"><span data-stu-id="5cbec-137">Example</span></span>

```javascript
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

####  <a name="officetheme-object"></a><span data-ttu-id="5cbec-138">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="5cbec-138">officeTheme :Object</span></span>

<span data-ttu-id="5cbec-139">提供对 Office 主题颜色的属性的访问权限。</span><span class="sxs-lookup"><span data-stu-id="5cbec-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="5cbec-140">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="5cbec-140">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="5cbec-p102">通过使用 Office 主题颜色，你可以使外接程序的配色方案与用户（通过 **“文件”>“Office 帐户”>“Office 主题”UI**）选择的当前 Office 主题协调一致，这种做法适用于所有 Office 主机应用程序。使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="5cbec-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="5cbec-143">类型</span><span class="sxs-lookup"><span data-stu-id="5cbec-143">Type</span></span>

*   <span data-ttu-id="5cbec-144">对象</span><span class="sxs-lookup"><span data-stu-id="5cbec-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="5cbec-145">属性：</span><span class="sxs-lookup"><span data-stu-id="5cbec-145">Properties:</span></span>

|<span data-ttu-id="5cbec-146">名称</span><span class="sxs-lookup"><span data-stu-id="5cbec-146">Name</span></span>| <span data-ttu-id="5cbec-147">类型</span><span class="sxs-lookup"><span data-stu-id="5cbec-147">Type</span></span>| <span data-ttu-id="5cbec-148">说明</span><span class="sxs-lookup"><span data-stu-id="5cbec-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="5cbec-149">String</span><span class="sxs-lookup"><span data-stu-id="5cbec-149">String</span></span>|<span data-ttu-id="5cbec-150">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="5cbec-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="5cbec-151">String</span><span class="sxs-lookup"><span data-stu-id="5cbec-151">String</span></span>|<span data-ttu-id="5cbec-152">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="5cbec-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="5cbec-153">字符串</span><span class="sxs-lookup"><span data-stu-id="5cbec-153">String</span></span>|<span data-ttu-id="5cbec-154">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="5cbec-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="5cbec-155">字符串</span><span class="sxs-lookup"><span data-stu-id="5cbec-155">String</span></span>|<span data-ttu-id="5cbec-156">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="5cbec-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5cbec-157">要求</span><span class="sxs-lookup"><span data-stu-id="5cbec-157">Requirements</span></span>

|<span data-ttu-id="5cbec-158">要求</span><span class="sxs-lookup"><span data-stu-id="5cbec-158">Requirement</span></span>| <span data-ttu-id="5cbec-159">值</span><span class="sxs-lookup"><span data-stu-id="5cbec-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cbec-160">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5cbec-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5cbec-161">1.3</span><span class="sxs-lookup"><span data-stu-id="5cbec-161">1.3</span></span>|
|[<span data-ttu-id="5cbec-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5cbec-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5cbec-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5cbec-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5cbec-164">示例</span><span class="sxs-lookup"><span data-stu-id="5cbec-164">Example</span></span>

```javascript
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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a><span data-ttu-id="5cbec-165">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="5cbec-165">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span></span>

<span data-ttu-id="5cbec-166">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="5cbec-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="5cbec-167">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="5cbec-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="5cbec-168">类型</span><span class="sxs-lookup"><span data-stu-id="5cbec-168">Type</span></span>

*   [<span data-ttu-id="5cbec-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="5cbec-169">RoamingSettings</span></span>](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="5cbec-170">要求</span><span class="sxs-lookup"><span data-stu-id="5cbec-170">Requirements</span></span>

|<span data-ttu-id="5cbec-171">要求</span><span class="sxs-lookup"><span data-stu-id="5cbec-171">Requirement</span></span>| <span data-ttu-id="5cbec-172">值</span><span class="sxs-lookup"><span data-stu-id="5cbec-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cbec-173">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="5cbec-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5cbec-174">1.0</span><span class="sxs-lookup"><span data-stu-id="5cbec-174">1.0</span></span>|
|[<span data-ttu-id="5cbec-175">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="5cbec-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5cbec-176">受限</span><span class="sxs-lookup"><span data-stu-id="5cbec-176">Restricted</span></span>|
|[<span data-ttu-id="5cbec-177">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="5cbec-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5cbec-178">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="5cbec-178">Compose or Read</span></span>|
