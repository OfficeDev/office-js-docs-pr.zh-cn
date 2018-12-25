---
title: Office.context - 要求集 1.1
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 392e54f1004bb395672c026ef749113f94ec7479
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432723"
---
# <a name="context"></a><span data-ttu-id="13831-102">context</span><span class="sxs-lookup"><span data-stu-id="13831-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="13831-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="13831-103">[Office](Office.md).context</span></span>

<span data-ttu-id="13831-p101">Office.context 命名空间提供所有 Office 应用中的外接程序所使用的共享接口。此列表仅记录 Outlook 外接程序所使用的接口。有关 Office.context 命名空间的完整列表，请参阅[共享 API 中的 Office.context 引用](/javascript/api/office/office.context)。</span><span class="sxs-lookup"><span data-stu-id="13831-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="13831-106">要求</span><span class="sxs-lookup"><span data-stu-id="13831-106">Requirements</span></span>

|<span data-ttu-id="13831-107">要求</span><span class="sxs-lookup"><span data-stu-id="13831-107">Requirement</span></span>| <span data-ttu-id="13831-108">值</span><span class="sxs-lookup"><span data-stu-id="13831-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="13831-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13831-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13831-110">1.0</span><span class="sxs-lookup"><span data-stu-id="13831-110">1.0</span></span>|
|[<span data-ttu-id="13831-111">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13831-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13831-112">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13831-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="13831-113">命名空间</span><span class="sxs-lookup"><span data-stu-id="13831-113">Namespaces</span></span>

<span data-ttu-id="13831-114">[mailbox](office.context.mailbox.md)：为 Microsoft Outlook 和 Microsoft Outlook 网页版提供对 Outlook 外接程序对象模型的访问权限。</span><span class="sxs-lookup"><span data-stu-id="13831-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="13831-115">成员</span><span class="sxs-lookup"><span data-stu-id="13831-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="13831-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="13831-116">displayLanguage :String</span></span>

<span data-ttu-id="13831-117">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="13831-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="13831-118">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="13831-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="13831-119">类型：</span><span class="sxs-lookup"><span data-stu-id="13831-119">Type:</span></span>

*   <span data-ttu-id="13831-120">String</span><span class="sxs-lookup"><span data-stu-id="13831-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="13831-121">要求</span><span class="sxs-lookup"><span data-stu-id="13831-121">Requirements</span></span>

|<span data-ttu-id="13831-122">要求</span><span class="sxs-lookup"><span data-stu-id="13831-122">Requirement</span></span>| <span data-ttu-id="13831-123">值</span><span class="sxs-lookup"><span data-stu-id="13831-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="13831-124">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13831-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13831-125">1.0</span><span class="sxs-lookup"><span data-stu-id="13831-125">1.0</span></span>|
|[<span data-ttu-id="13831-126">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13831-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13831-127">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13831-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13831-128">示例</span><span class="sxs-lookup"><span data-stu-id="13831-128">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a><span data-ttu-id="13831-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="13831-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span></span>

<span data-ttu-id="13831-130">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="13831-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="13831-131">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="13831-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="13831-132">类型:</span><span class="sxs-lookup"><span data-stu-id="13831-132">Type:</span></span>

*   [<span data-ttu-id="13831-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="13831-133">RoamingSettings</span></span>](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="13831-134">要求</span><span class="sxs-lookup"><span data-stu-id="13831-134">Requirements</span></span>

|<span data-ttu-id="13831-135">要求</span><span class="sxs-lookup"><span data-stu-id="13831-135">Requirement</span></span>| <span data-ttu-id="13831-136">值</span><span class="sxs-lookup"><span data-stu-id="13831-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="13831-137">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13831-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13831-138">1.0</span><span class="sxs-lookup"><span data-stu-id="13831-138">1.0</span></span>|
|[<span data-ttu-id="13831-139">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13831-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13831-140">受限</span><span class="sxs-lookup"><span data-stu-id="13831-140">Restricted</span></span>|
|[<span data-ttu-id="13831-141">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13831-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13831-142">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13831-142">Compose or read</span></span>|