---
title: Office。上下文要求集1。7
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 8f21389d40181db0f62a8e13a882aaf4cb9b5072
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454858"
---
# <a name="context"></a><span data-ttu-id="db18a-102">context</span><span class="sxs-lookup"><span data-stu-id="db18a-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="db18a-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="db18a-103">[Office](Office.md).context</span></span>

<span data-ttu-id="db18a-p101">Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[通用 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="db18a-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="db18a-106">要求</span><span class="sxs-lookup"><span data-stu-id="db18a-106">Requirements</span></span>

|<span data-ttu-id="db18a-107">要求</span><span class="sxs-lookup"><span data-stu-id="db18a-107">Requirement</span></span>| <span data-ttu-id="db18a-108">值</span><span class="sxs-lookup"><span data-stu-id="db18a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="db18a-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="db18a-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="db18a-110">1.0</span><span class="sxs-lookup"><span data-stu-id="db18a-110">1.0</span></span>|
|[<span data-ttu-id="db18a-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="db18a-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="db18a-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="db18a-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="db18a-113">成员和方法</span><span class="sxs-lookup"><span data-stu-id="db18a-113">Members and methods</span></span>

| <span data-ttu-id="db18a-114">成员</span><span class="sxs-lookup"><span data-stu-id="db18a-114">Member</span></span> | <span data-ttu-id="db18a-115">类型</span><span class="sxs-lookup"><span data-stu-id="db18a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="db18a-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="db18a-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="db18a-117">Member</span><span class="sxs-lookup"><span data-stu-id="db18a-117">Member</span></span> |
| [<span data-ttu-id="db18a-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="db18a-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="db18a-119">成员</span><span class="sxs-lookup"><span data-stu-id="db18a-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="db18a-120">命名空间</span><span class="sxs-lookup"><span data-stu-id="db18a-120">Namespaces</span></span>

<span data-ttu-id="db18a-121">[邮箱](office.context.mailbox.md): 提供对 Microsoft Outlook 的 outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="db18a-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="db18a-122">Members</span><span class="sxs-lookup"><span data-stu-id="db18a-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="db18a-123">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="db18a-123">displayLanguage: String</span></span>

<span data-ttu-id="db18a-124">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="db18a-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="db18a-125">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="db18a-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="db18a-126">类型</span><span class="sxs-lookup"><span data-stu-id="db18a-126">Type</span></span>

*   <span data-ttu-id="db18a-127">String</span><span class="sxs-lookup"><span data-stu-id="db18a-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="db18a-128">要求</span><span class="sxs-lookup"><span data-stu-id="db18a-128">Requirements</span></span>

|<span data-ttu-id="db18a-129">要求</span><span class="sxs-lookup"><span data-stu-id="db18a-129">Requirement</span></span>| <span data-ttu-id="db18a-130">值</span><span class="sxs-lookup"><span data-stu-id="db18a-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="db18a-131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="db18a-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="db18a-132">1.0</span><span class="sxs-lookup"><span data-stu-id="db18a-132">1.0</span></span>|
|[<span data-ttu-id="db18a-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="db18a-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="db18a-134">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="db18a-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="db18a-135">示例</span><span class="sxs-lookup"><span data-stu-id="db18a-135">Example</span></span>

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

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a><span data-ttu-id="db18a-136">roamingSettings: [roamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="db18a-136">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span></span>

<span data-ttu-id="db18a-137">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="db18a-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="db18a-138">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="db18a-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="db18a-139">类型</span><span class="sxs-lookup"><span data-stu-id="db18a-139">Type</span></span>

*   [<span data-ttu-id="db18a-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="db18a-140">RoamingSettings</span></span>](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="db18a-141">要求</span><span class="sxs-lookup"><span data-stu-id="db18a-141">Requirements</span></span>

|<span data-ttu-id="db18a-142">要求</span><span class="sxs-lookup"><span data-stu-id="db18a-142">Requirement</span></span>| <span data-ttu-id="db18a-143">值</span><span class="sxs-lookup"><span data-stu-id="db18a-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="db18a-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="db18a-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="db18a-145">1.0</span><span class="sxs-lookup"><span data-stu-id="db18a-145">1.0</span></span>|
|[<span data-ttu-id="db18a-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="db18a-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="db18a-147">受限</span><span class="sxs-lookup"><span data-stu-id="db18a-147">Restricted</span></span>|
|[<span data-ttu-id="db18a-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="db18a-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="db18a-149">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="db18a-149">Compose or Read</span></span>|
