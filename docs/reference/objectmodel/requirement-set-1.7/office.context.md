---
title: Office。上下文要求集1。7
description: 使用邮箱 API 要求集1.7 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 195b9e9a6aa21ffbc356a628291587c69d63326b
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293700"
---
# <a name="context-mailbox-requirement-set-17"></a><span data-ttu-id="8f6d0-103"> (邮箱要求集1.7 的上下文) </span><span class="sxs-lookup"><span data-stu-id="8f6d0-103">context (Mailbox requirement set 1.7)</span></span>

### <a name="officecontext"></a><span data-ttu-id="8f6d0-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="8f6d0-104">[Office](office.md).context</span></span>

<span data-ttu-id="8f6d0-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="8f6d0-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.7)"。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.7).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f6d0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="8f6d0-107">Requirements</span></span>

|<span data-ttu-id="8f6d0-108">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-108">Requirement</span></span>| <span data-ttu-id="8f6d0-109">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-111">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-111">1.1</span></span>|
|[<span data-ttu-id="8f6d0-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8f6d0-114">属性</span><span class="sxs-lookup"><span data-stu-id="8f6d0-114">Properties</span></span>

| <span data-ttu-id="8f6d0-115">属性</span><span class="sxs-lookup"><span data-stu-id="8f6d0-115">Property</span></span> | <span data-ttu-id="8f6d0-116">型号</span><span class="sxs-lookup"><span data-stu-id="8f6d0-116">Modes</span></span> | <span data-ttu-id="8f6d0-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-117">Return type</span></span> | <span data-ttu-id="8f6d0-118">最小值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-118">Minimum</span></span><br><span data-ttu-id="8f6d0-119">要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8f6d0-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="8f6d0-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="8f6d0-121">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-121">Compose</span></span><br><span data-ttu-id="8f6d0-122">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-122">Read</span></span> | <span data-ttu-id="8f6d0-123">String</span><span class="sxs-lookup"><span data-stu-id="8f6d0-123">String</span></span> | [<span data-ttu-id="8f6d0-124">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-125">过程</span><span class="sxs-lookup"><span data-stu-id="8f6d0-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="8f6d0-126">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-126">Compose</span></span><br><span data-ttu-id="8f6d0-127">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-127">Read</span></span> | [<span data-ttu-id="8f6d0-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8f6d0-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.7) | [<span data-ttu-id="8f6d0-129">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="8f6d0-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="8f6d0-131">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-131">Compose</span></span><br><span data-ttu-id="8f6d0-132">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-132">Read</span></span> | <span data-ttu-id="8f6d0-133">String</span><span class="sxs-lookup"><span data-stu-id="8f6d0-133">String</span></span> | [<span data-ttu-id="8f6d0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-135">host</span><span class="sxs-lookup"><span data-stu-id="8f6d0-135">host</span></span>](#host-hosttype) | <span data-ttu-id="8f6d0-136">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-136">Compose</span></span><br><span data-ttu-id="8f6d0-137">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-137">Read</span></span> | [<span data-ttu-id="8f6d0-138">HostType</span><span class="sxs-lookup"><span data-stu-id="8f6d0-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.7) | [<span data-ttu-id="8f6d0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="8f6d0-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="8f6d0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-141">Compose</span></span><br><span data-ttu-id="8f6d0-142">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-142">Read</span></span> | [<span data-ttu-id="8f6d0-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="8f6d0-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7) | [<span data-ttu-id="8f6d0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-145">平台</span><span class="sxs-lookup"><span data-stu-id="8f6d0-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="8f6d0-146">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-146">Compose</span></span><br><span data-ttu-id="8f6d0-147">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-147">Read</span></span> | [<span data-ttu-id="8f6d0-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8f6d0-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.7) | [<span data-ttu-id="8f6d0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-150">满足</span><span class="sxs-lookup"><span data-stu-id="8f6d0-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="8f6d0-151">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-151">Compose</span></span><br><span data-ttu-id="8f6d0-152">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-152">Read</span></span> | [<span data-ttu-id="8f6d0-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8f6d0-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7) | [<span data-ttu-id="8f6d0-154">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="8f6d0-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="8f6d0-156">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-156">Compose</span></span><br><span data-ttu-id="8f6d0-157">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-157">Read</span></span> | [<span data-ttu-id="8f6d0-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8f6d0-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7) | [<span data-ttu-id="8f6d0-159">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f6d0-160">ui</span><span class="sxs-lookup"><span data-stu-id="8f6d0-160">ui</span></span>](#ui-ui) | <span data-ttu-id="8f6d0-161">撰写</span><span class="sxs-lookup"><span data-stu-id="8f6d0-161">Compose</span></span><br><span data-ttu-id="8f6d0-162">阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-162">Read</span></span> | [<span data-ttu-id="8f6d0-163">UI</span><span class="sxs-lookup"><span data-stu-id="8f6d0-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.7) | [<span data-ttu-id="8f6d0-164">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="8f6d0-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="8f6d0-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="8f6d0-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="8f6d0-166">contentLanguage: String</span></span>

<span data-ttu-id="8f6d0-167">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="8f6d0-168">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-169">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-169">Type</span></span>

*   <span data-ttu-id="8f6d0-170">String</span><span class="sxs-lookup"><span data-stu-id="8f6d0-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f6d0-171">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-171">Requirements</span></span>

|<span data-ttu-id="8f6d0-172">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-172">Requirement</span></span>| <span data-ttu-id="8f6d0-173">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-175">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-175">1.1</span></span>|
|[<span data-ttu-id="8f6d0-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f6d0-178">示例</span><span class="sxs-lookup"><span data-stu-id="8f6d0-178">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
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

<br>

---
---

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="8f6d0-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="8f6d0-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="8f6d0-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-181">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-181">Type</span></span>

*   [<span data-ttu-id="8f6d0-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8f6d0-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="8f6d0-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="8f6d0-183">Requirements</span></span>

|<span data-ttu-id="8f6d0-184">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-184">Requirement</span></span>| <span data-ttu-id="8f6d0-185">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-187">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-187">1.1</span></span>|
|[<span data-ttu-id="8f6d0-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f6d0-190">示例</span><span class="sxs-lookup"><span data-stu-id="8f6d0-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="8f6d0-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="8f6d0-191">displayLanguage: String</span></span>

<span data-ttu-id="8f6d0-192">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="8f6d0-193">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的**File > Options > 语言**指定的当前**显示语言**设置。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-194">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-194">Type</span></span>

*   <span data-ttu-id="8f6d0-195">String</span><span class="sxs-lookup"><span data-stu-id="8f6d0-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f6d0-196">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-196">Requirements</span></span>

|<span data-ttu-id="8f6d0-197">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-197">Requirement</span></span>| <span data-ttu-id="8f6d0-198">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-200">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-200">1.1</span></span>|
|[<span data-ttu-id="8f6d0-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f6d0-203">示例</span><span class="sxs-lookup"><span data-stu-id="8f6d0-203">Example</span></span>

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

<br>

---
---

#### <a name="host-hosttype"></a><span data-ttu-id="8f6d0-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="8f6d0-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="8f6d0-205">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-206">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-206">Type</span></span>

*   [<span data-ttu-id="8f6d0-207">HostType</span><span class="sxs-lookup"><span data-stu-id="8f6d0-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="8f6d0-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="8f6d0-208">Requirements</span></span>

|<span data-ttu-id="8f6d0-209">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-209">Requirement</span></span>| <span data-ttu-id="8f6d0-210">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-212">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-212">1.1</span></span>|
|[<span data-ttu-id="8f6d0-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-214">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f6d0-215">示例</span><span class="sxs-lookup"><span data-stu-id="8f6d0-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="8f6d0-216">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="8f6d0-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="8f6d0-217">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-218">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-218">Type</span></span>

*   [<span data-ttu-id="8f6d0-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8f6d0-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="8f6d0-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="8f6d0-220">Requirements</span></span>

|<span data-ttu-id="8f6d0-221">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-221">Requirement</span></span>| <span data-ttu-id="8f6d0-222">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-224">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-224">1.1</span></span>|
|[<span data-ttu-id="8f6d0-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f6d0-227">示例</span><span class="sxs-lookup"><span data-stu-id="8f6d0-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="8f6d0-228">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="8f6d0-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="8f6d0-229">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-230">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-230">Type</span></span>

*   [<span data-ttu-id="8f6d0-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8f6d0-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="8f6d0-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="8f6d0-232">Requirements</span></span>

|<span data-ttu-id="8f6d0-233">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-233">Requirement</span></span>| <span data-ttu-id="8f6d0-234">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-236">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-236">1.1</span></span>|
|[<span data-ttu-id="8f6d0-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f6d0-239">示例</span><span class="sxs-lookup"><span data-stu-id="8f6d0-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="8f6d0-240">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="8f6d0-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="8f6d0-241">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8f6d0-242">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-243">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-243">Type</span></span>

*   [<span data-ttu-id="8f6d0-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8f6d0-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8f6d0-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="8f6d0-245">Requirements</span></span>

|<span data-ttu-id="8f6d0-246">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-246">Requirement</span></span>| <span data-ttu-id="8f6d0-247">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-249">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-249">1.1</span></span>|
|[<span data-ttu-id="8f6d0-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8f6d0-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="8f6d0-251">受限</span><span class="sxs-lookup"><span data-stu-id="8f6d0-251">Restricted</span></span>|
|[<span data-ttu-id="8f6d0-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="8f6d0-254">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="8f6d0-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="8f6d0-255">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="8f6d0-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8f6d0-256">类型</span><span class="sxs-lookup"><span data-stu-id="8f6d0-256">Type</span></span>

*   [<span data-ttu-id="8f6d0-257">UI</span><span class="sxs-lookup"><span data-stu-id="8f6d0-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="8f6d0-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="8f6d0-258">Requirements</span></span>

|<span data-ttu-id="8f6d0-259">要求</span><span class="sxs-lookup"><span data-stu-id="8f6d0-259">Requirement</span></span>| <span data-ttu-id="8f6d0-260">值</span><span class="sxs-lookup"><span data-stu-id="8f6d0-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f6d0-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8f6d0-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f6d0-262">1.1</span><span class="sxs-lookup"><span data-stu-id="8f6d0-262">1.1</span></span>|
|[<span data-ttu-id="8f6d0-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8f6d0-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f6d0-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8f6d0-264">Compose or Read</span></span>|
