---
title: Office.context - 要求集 1.3
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 78fc4aac0baf8126599eea86a1c705f9d6311b59
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432786"
---
# <a name="context"></a><span data-ttu-id="6050e-102">context</span><span class="sxs-lookup"><span data-stu-id="6050e-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="6050e-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="6050e-103">[Office](Office.md).context</span></span>

<span data-ttu-id="6050e-p101">Office.context 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[共享 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="6050e-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6050e-106">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-106">Requirements</span></span>

|<span data-ttu-id="6050e-107">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-107">Requirement</span></span>| <span data-ttu-id="6050e-108">值</span><span class="sxs-lookup"><span data-stu-id="6050e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="6050e-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6050e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6050e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="6050e-110">1.0</span></span>|
|[<span data-ttu-id="6050e-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6050e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6050e-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6050e-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="6050e-113">命名空间</span><span class="sxs-lookup"><span data-stu-id="6050e-113">Namespaces</span></span>

<span data-ttu-id="6050e-114">[mailbox](office.context.mailbox.md)：为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="6050e-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="6050e-115">成员</span><span class="sxs-lookup"><span data-stu-id="6050e-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="6050e-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="6050e-116">displayLanguage :String</span></span>

<span data-ttu-id="6050e-117">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="6050e-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="6050e-118">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="6050e-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="6050e-119">类型：</span><span class="sxs-lookup"><span data-stu-id="6050e-119">Type:</span></span>

*   <span data-ttu-id="6050e-120">String</span><span class="sxs-lookup"><span data-stu-id="6050e-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6050e-121">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-121">Requirements</span></span>

|<span data-ttu-id="6050e-122">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-122">Requirement</span></span>| <span data-ttu-id="6050e-123">值</span><span class="sxs-lookup"><span data-stu-id="6050e-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="6050e-124">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6050e-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6050e-125">1.0</span><span class="sxs-lookup"><span data-stu-id="6050e-125">1.0</span></span>|
|[<span data-ttu-id="6050e-126">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6050e-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6050e-127">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6050e-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6050e-128">示例</span><span class="sxs-lookup"><span data-stu-id="6050e-128">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="6050e-129">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="6050e-129">officeTheme :Object</span></span>

<span data-ttu-id="6050e-130">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="6050e-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="6050e-131">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="6050e-131">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6050e-p102">通过使用 Office 主题颜色，你可以使外接程序的配色方案与用户（通过 **“文件”>“Office 帐户”>“Office 主题”UI**）选择的当前 Office 主题协调一致，这种做法适用于所有 Office 主机应用程序。使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="6050e-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6050e-134">类型：</span><span class="sxs-lookup"><span data-stu-id="6050e-134">Type:</span></span>

*   <span data-ttu-id="6050e-135">对象</span><span class="sxs-lookup"><span data-stu-id="6050e-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="6050e-136">属性：</span><span class="sxs-lookup"><span data-stu-id="6050e-136">Properties:</span></span>

|<span data-ttu-id="6050e-137">名称</span><span class="sxs-lookup"><span data-stu-id="6050e-137">Name</span></span>| <span data-ttu-id="6050e-138">类型</span><span class="sxs-lookup"><span data-stu-id="6050e-138">Type</span></span>| <span data-ttu-id="6050e-139">描述</span><span class="sxs-lookup"><span data-stu-id="6050e-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="6050e-140">String</span><span class="sxs-lookup"><span data-stu-id="6050e-140">String</span></span>|<span data-ttu-id="6050e-141">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="6050e-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="6050e-142">String</span><span class="sxs-lookup"><span data-stu-id="6050e-142">String</span></span>|<span data-ttu-id="6050e-143">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="6050e-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="6050e-144">字符串</span><span class="sxs-lookup"><span data-stu-id="6050e-144">String</span></span>|<span data-ttu-id="6050e-145">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="6050e-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="6050e-146">字符串</span><span class="sxs-lookup"><span data-stu-id="6050e-146">String</span></span>|<span data-ttu-id="6050e-147">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="6050e-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6050e-148">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-148">Requirements</span></span>

|<span data-ttu-id="6050e-149">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-149">Requirement</span></span>| <span data-ttu-id="6050e-150">值</span><span class="sxs-lookup"><span data-stu-id="6050e-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="6050e-151">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6050e-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6050e-152">1.3</span><span class="sxs-lookup"><span data-stu-id="6050e-152">1.3</span></span>|
|[<span data-ttu-id="6050e-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6050e-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6050e-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6050e-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6050e-155">示例</span><span class="sxs-lookup"><span data-stu-id="6050e-155">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="6050e-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="6050e-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="6050e-157">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="6050e-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6050e-158">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="6050e-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6050e-159">类型:</span><span class="sxs-lookup"><span data-stu-id="6050e-159">Type:</span></span>

*   [<span data-ttu-id="6050e-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6050e-160">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6050e-161">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-161">Requirements</span></span>

|<span data-ttu-id="6050e-162">要求</span><span class="sxs-lookup"><span data-stu-id="6050e-162">Requirement</span></span>| <span data-ttu-id="6050e-163">值</span><span class="sxs-lookup"><span data-stu-id="6050e-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="6050e-164">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6050e-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6050e-165">1.0</span><span class="sxs-lookup"><span data-stu-id="6050e-165">1.0</span></span>|
|[<span data-ttu-id="6050e-166">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6050e-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6050e-167">受限</span><span class="sxs-lookup"><span data-stu-id="6050e-167">Restricted</span></span>|
|[<span data-ttu-id="6050e-168">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6050e-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6050e-169">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6050e-169">Compose or read</span></span>|