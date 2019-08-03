---
title: Office。上下文要求集1。4
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 7f4637a1d6a4a9bc2f97d039ed4404ab549a2b34
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064646"
---
# <a name="context"></a><span data-ttu-id="04687-102">context</span><span class="sxs-lookup"><span data-stu-id="04687-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="04687-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="04687-103">[Office](Office.md).context</span></span>

<span data-ttu-id="04687-p101">Office.context 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[通用 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="04687-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="04687-106">要求</span><span class="sxs-lookup"><span data-stu-id="04687-106">Requirements</span></span>

|<span data-ttu-id="04687-107">要求</span><span class="sxs-lookup"><span data-stu-id="04687-107">Requirement</span></span>| <span data-ttu-id="04687-108">值</span><span class="sxs-lookup"><span data-stu-id="04687-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="04687-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="04687-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04687-110">1.0</span><span class="sxs-lookup"><span data-stu-id="04687-110">1.0</span></span>|
|[<span data-ttu-id="04687-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="04687-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04687-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="04687-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="04687-113">命名空间</span><span class="sxs-lookup"><span data-stu-id="04687-113">Namespaces</span></span>

<span data-ttu-id="04687-114">[邮箱](office.context.mailbox.md): 提供对 Microsoft Outlook 的 outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="04687-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="04687-115">Members</span><span class="sxs-lookup"><span data-stu-id="04687-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="04687-116">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="04687-116">displayLanguage: String</span></span>

<span data-ttu-id="04687-117">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="04687-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="04687-118">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="04687-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="04687-119">类型</span><span class="sxs-lookup"><span data-stu-id="04687-119">Type</span></span>

*   <span data-ttu-id="04687-120">String</span><span class="sxs-lookup"><span data-stu-id="04687-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="04687-121">要求</span><span class="sxs-lookup"><span data-stu-id="04687-121">Requirements</span></span>

|<span data-ttu-id="04687-122">要求</span><span class="sxs-lookup"><span data-stu-id="04687-122">Requirement</span></span>| <span data-ttu-id="04687-123">值</span><span class="sxs-lookup"><span data-stu-id="04687-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="04687-124">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="04687-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04687-125">1.0</span><span class="sxs-lookup"><span data-stu-id="04687-125">1.0</span></span>|
|[<span data-ttu-id="04687-126">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="04687-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04687-127">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="04687-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="04687-128">示例</span><span class="sxs-lookup"><span data-stu-id="04687-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-14"></a><span data-ttu-id="04687-129">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="04687-129">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)</span></span>

<span data-ttu-id="04687-130">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="04687-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="04687-131">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="04687-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="04687-132">类型</span><span class="sxs-lookup"><span data-stu-id="04687-132">Type</span></span>

*   [<span data-ttu-id="04687-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="04687-133">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="04687-134">要求</span><span class="sxs-lookup"><span data-stu-id="04687-134">Requirements</span></span>

|<span data-ttu-id="04687-135">要求</span><span class="sxs-lookup"><span data-stu-id="04687-135">Requirement</span></span>| <span data-ttu-id="04687-136">值</span><span class="sxs-lookup"><span data-stu-id="04687-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="04687-137">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="04687-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04687-138">1.0</span><span class="sxs-lookup"><span data-stu-id="04687-138">1.0</span></span>|
|[<span data-ttu-id="04687-139">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="04687-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="04687-140">受限</span><span class="sxs-lookup"><span data-stu-id="04687-140">Restricted</span></span>|
|[<span data-ttu-id="04687-141">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="04687-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="04687-142">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="04687-142">Compose or Read</span></span>|
