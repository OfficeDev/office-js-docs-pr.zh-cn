---
title: Office。上下文要求集1。4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 738c6a5ffbe6bb59f77e3bb82baee78a40be136e
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268308"
---
# <a name="context"></a><span data-ttu-id="ae321-102">context</span><span class="sxs-lookup"><span data-stu-id="ae321-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="ae321-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="ae321-103">[Office](Office.md).context</span></span>

<span data-ttu-id="ae321-p101">Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[通用 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="ae321-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae321-106">要求</span><span class="sxs-lookup"><span data-stu-id="ae321-106">Requirements</span></span>

|<span data-ttu-id="ae321-107">要求</span><span class="sxs-lookup"><span data-stu-id="ae321-107">Requirement</span></span>| <span data-ttu-id="ae321-108">值</span><span class="sxs-lookup"><span data-stu-id="ae321-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae321-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae321-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae321-110">1.0</span><span class="sxs-lookup"><span data-stu-id="ae321-110">1.0</span></span>|
|[<span data-ttu-id="ae321-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae321-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae321-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae321-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ae321-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="ae321-113">Members and methods</span></span>

| <span data-ttu-id="ae321-114">成员</span><span class="sxs-lookup"><span data-stu-id="ae321-114">Member</span></span> | <span data-ttu-id="ae321-115">类型</span><span class="sxs-lookup"><span data-stu-id="ae321-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ae321-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ae321-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ae321-117">Member</span><span class="sxs-lookup"><span data-stu-id="ae321-117">Member</span></span> |
| [<span data-ttu-id="ae321-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae321-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ae321-119">成员</span><span class="sxs-lookup"><span data-stu-id="ae321-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ae321-120">命名空间</span><span class="sxs-lookup"><span data-stu-id="ae321-120">Namespaces</span></span>

<span data-ttu-id="ae321-121">[邮箱](office.context.mailbox.md): 提供对 Microsoft Outlook 的 outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="ae321-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="ae321-122">Members</span><span class="sxs-lookup"><span data-stu-id="ae321-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="ae321-123">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="ae321-123">displayLanguage: String</span></span>

<span data-ttu-id="ae321-124">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="ae321-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="ae321-125">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="ae321-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ae321-126">类型</span><span class="sxs-lookup"><span data-stu-id="ae321-126">Type</span></span>

*   <span data-ttu-id="ae321-127">String</span><span class="sxs-lookup"><span data-stu-id="ae321-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae321-128">要求</span><span class="sxs-lookup"><span data-stu-id="ae321-128">Requirements</span></span>

|<span data-ttu-id="ae321-129">要求</span><span class="sxs-lookup"><span data-stu-id="ae321-129">Requirement</span></span>| <span data-ttu-id="ae321-130">值</span><span class="sxs-lookup"><span data-stu-id="ae321-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae321-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae321-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae321-132">1.0</span><span class="sxs-lookup"><span data-stu-id="ae321-132">1.0</span></span>|
|[<span data-ttu-id="ae321-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae321-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae321-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae321-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae321-135">示例</span><span class="sxs-lookup"><span data-stu-id="ae321-135">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-14"></a><span data-ttu-id="ae321-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="ae321-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)</span></span>

<span data-ttu-id="ae321-137">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="ae321-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ae321-138">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="ae321-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ae321-139">类型</span><span class="sxs-lookup"><span data-stu-id="ae321-139">Type</span></span>

*   [<span data-ttu-id="ae321-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae321-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="ae321-141">要求</span><span class="sxs-lookup"><span data-stu-id="ae321-141">Requirements</span></span>

|<span data-ttu-id="ae321-142">要求</span><span class="sxs-lookup"><span data-stu-id="ae321-142">Requirement</span></span>| <span data-ttu-id="ae321-143">值</span><span class="sxs-lookup"><span data-stu-id="ae321-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae321-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae321-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae321-145">1.0</span><span class="sxs-lookup"><span data-stu-id="ae321-145">1.0</span></span>|
|[<span data-ttu-id="ae321-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae321-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae321-147">受限</span><span class="sxs-lookup"><span data-stu-id="ae321-147">Restricted</span></span>|
|[<span data-ttu-id="ae321-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae321-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae321-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae321-149">Compose or Read</span></span>|
