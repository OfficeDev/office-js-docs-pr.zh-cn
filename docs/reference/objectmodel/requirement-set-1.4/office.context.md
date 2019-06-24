---
title: Office。上下文要求集1。4
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: ad1887f32568f30cb87e52dd1f9457be2022beb2
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127336"
---
# <a name="context"></a><span data-ttu-id="e49ea-102">context</span><span class="sxs-lookup"><span data-stu-id="e49ea-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="e49ea-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="e49ea-103">[Office](Office.md).context</span></span>

<span data-ttu-id="e49ea-p101">Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[通用 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="e49ea-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e49ea-106">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-106">Requirements</span></span>

|<span data-ttu-id="e49ea-107">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-107">Requirement</span></span>| <span data-ttu-id="e49ea-108">值</span><span class="sxs-lookup"><span data-stu-id="e49ea-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e49ea-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e49ea-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e49ea-110">1.0</span><span class="sxs-lookup"><span data-stu-id="e49ea-110">1.0</span></span>|
|[<span data-ttu-id="e49ea-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e49ea-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e49ea-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e49ea-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="e49ea-113">命名空间</span><span class="sxs-lookup"><span data-stu-id="e49ea-113">Namespaces</span></span>

<span data-ttu-id="e49ea-114">[邮箱](office.context.mailbox.md): 提供对 Microsoft Outlook 的 outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e49ea-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="e49ea-115">Members</span><span class="sxs-lookup"><span data-stu-id="e49ea-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="e49ea-116">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="e49ea-116">displayLanguage: String</span></span>

<span data-ttu-id="e49ea-117">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="e49ea-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="e49ea-118">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="e49ea-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="e49ea-119">类型</span><span class="sxs-lookup"><span data-stu-id="e49ea-119">Type</span></span>

*   <span data-ttu-id="e49ea-120">String</span><span class="sxs-lookup"><span data-stu-id="e49ea-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e49ea-121">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-121">Requirements</span></span>

|<span data-ttu-id="e49ea-122">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-122">Requirement</span></span>| <span data-ttu-id="e49ea-123">值</span><span class="sxs-lookup"><span data-stu-id="e49ea-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="e49ea-124">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e49ea-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e49ea-125">1.0</span><span class="sxs-lookup"><span data-stu-id="e49ea-125">1.0</span></span>|
|[<span data-ttu-id="e49ea-126">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e49ea-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e49ea-127">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e49ea-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e49ea-128">示例</span><span class="sxs-lookup"><span data-stu-id="e49ea-128">Example</span></span>

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

#### <a name="officetheme-object"></a><span data-ttu-id="e49ea-129">officeTheme: Object</span><span class="sxs-lookup"><span data-stu-id="e49ea-129">officeTheme: Object</span></span>

<span data-ttu-id="e49ea-130">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="e49ea-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="e49ea-131">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="e49ea-131">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e49ea-p102">通过使用 Office 主题颜色，你可以使外接程序的配色方案与用户（通过 **“文件”>“Office 帐户”>“Office 主题”UI**）选择的当前 Office 主题协调一致，这种做法适用于所有 Office 主机应用程序。使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="e49ea-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e49ea-134">类型</span><span class="sxs-lookup"><span data-stu-id="e49ea-134">Type</span></span>

*   <span data-ttu-id="e49ea-135">对象</span><span class="sxs-lookup"><span data-stu-id="e49ea-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="e49ea-136">属性：</span><span class="sxs-lookup"><span data-stu-id="e49ea-136">Properties:</span></span>

|<span data-ttu-id="e49ea-137">名称</span><span class="sxs-lookup"><span data-stu-id="e49ea-137">Name</span></span>| <span data-ttu-id="e49ea-138">类型</span><span class="sxs-lookup"><span data-stu-id="e49ea-138">Type</span></span>| <span data-ttu-id="e49ea-139">说明</span><span class="sxs-lookup"><span data-stu-id="e49ea-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="e49ea-140">String</span><span class="sxs-lookup"><span data-stu-id="e49ea-140">String</span></span>|<span data-ttu-id="e49ea-141">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="e49ea-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="e49ea-142">String</span><span class="sxs-lookup"><span data-stu-id="e49ea-142">String</span></span>|<span data-ttu-id="e49ea-143">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="e49ea-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="e49ea-144">String</span><span class="sxs-lookup"><span data-stu-id="e49ea-144">String</span></span>|<span data-ttu-id="e49ea-145">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="e49ea-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="e49ea-146">字符串</span><span class="sxs-lookup"><span data-stu-id="e49ea-146">String</span></span>|<span data-ttu-id="e49ea-147">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="e49ea-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e49ea-148">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-148">Requirements</span></span>

|<span data-ttu-id="e49ea-149">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-149">Requirement</span></span>| <span data-ttu-id="e49ea-150">值</span><span class="sxs-lookup"><span data-stu-id="e49ea-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="e49ea-151">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e49ea-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e49ea-152">1.3</span><span class="sxs-lookup"><span data-stu-id="e49ea-152">1.3</span></span>|
|[<span data-ttu-id="e49ea-153">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e49ea-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e49ea-154">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e49ea-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e49ea-155">示例</span><span class="sxs-lookup"><span data-stu-id="e49ea-155">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="e49ea-156">roamingSettings: [roamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="e49ea-156">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="e49ea-157">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="e49ea-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e49ea-158">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="e49ea-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e49ea-159">类型</span><span class="sxs-lookup"><span data-stu-id="e49ea-159">Type</span></span>

*   [<span data-ttu-id="e49ea-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e49ea-160">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e49ea-161">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-161">Requirements</span></span>

|<span data-ttu-id="e49ea-162">要求</span><span class="sxs-lookup"><span data-stu-id="e49ea-162">Requirement</span></span>| <span data-ttu-id="e49ea-163">值</span><span class="sxs-lookup"><span data-stu-id="e49ea-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="e49ea-164">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e49ea-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e49ea-165">1.0</span><span class="sxs-lookup"><span data-stu-id="e49ea-165">1.0</span></span>|
|[<span data-ttu-id="e49ea-166">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e49ea-166">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e49ea-167">受限</span><span class="sxs-lookup"><span data-stu-id="e49ea-167">Restricted</span></span>|
|[<span data-ttu-id="e49ea-168">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e49ea-168">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e49ea-169">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e49ea-169">Compose or Read</span></span>|
