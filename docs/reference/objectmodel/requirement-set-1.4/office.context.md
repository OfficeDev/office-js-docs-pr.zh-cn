---
title: Office。上下文要求集1。4
description: Outlook 外接程序 API 中的 Outlook 上下文对象的对象模型（邮箱 API 1.4 版本）。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: d2d5520d4b48700b2e232e0782583ebd5f2ce6c3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717724"
---
# <a name="context"></a><span data-ttu-id="1ad0d-103">context</span><span class="sxs-lookup"><span data-stu-id="1ad0d-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="1ad0d-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="1ad0d-104">[Office](office.md).context</span></span>

<span data-ttu-id="1ad0d-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="1ad0d-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.4)"。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ad0d-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ad0d-107">Requirements</span></span>

|<span data-ttu-id="1ad0d-108">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-108">Requirement</span></span>| <span data-ttu-id="1ad0d-109">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-111">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-111">1.1</span></span>|
|[<span data-ttu-id="1ad0d-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1ad0d-114">属性</span><span class="sxs-lookup"><span data-stu-id="1ad0d-114">Properties</span></span>

| <span data-ttu-id="1ad0d-115">属性</span><span class="sxs-lookup"><span data-stu-id="1ad0d-115">Property</span></span> | <span data-ttu-id="1ad0d-116">型号</span><span class="sxs-lookup"><span data-stu-id="1ad0d-116">Modes</span></span> | <span data-ttu-id="1ad0d-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-117">Return type</span></span> | <span data-ttu-id="1ad0d-118">最低</span><span class="sxs-lookup"><span data-stu-id="1ad0d-118">Minimum</span></span><br><span data-ttu-id="1ad0d-119">要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1ad0d-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="1ad0d-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="1ad0d-121">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-121">Compose</span></span><br><span data-ttu-id="1ad0d-122">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-122">Read</span></span> | <span data-ttu-id="1ad0d-123">String</span><span class="sxs-lookup"><span data-stu-id="1ad0d-123">String</span></span> | [<span data-ttu-id="1ad0d-124">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-125">过程</span><span class="sxs-lookup"><span data-stu-id="1ad0d-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="1ad0d-126">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-126">Compose</span></span><br><span data-ttu-id="1ad0d-127">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-127">Read</span></span> | [<span data-ttu-id="1ad0d-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1ad0d-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="1ad0d-129">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="1ad0d-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="1ad0d-131">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-131">Compose</span></span><br><span data-ttu-id="1ad0d-132">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-132">Read</span></span> | <span data-ttu-id="1ad0d-133">String</span><span class="sxs-lookup"><span data-stu-id="1ad0d-133">String</span></span> | [<span data-ttu-id="1ad0d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-135">host</span><span class="sxs-lookup"><span data-stu-id="1ad0d-135">host</span></span>](#host-hosttype) | <span data-ttu-id="1ad0d-136">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-136">Compose</span></span><br><span data-ttu-id="1ad0d-137">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-137">Read</span></span> | [<span data-ttu-id="1ad0d-138">HostType</span><span class="sxs-lookup"><span data-stu-id="1ad0d-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="1ad0d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="1ad0d-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="1ad0d-141">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-141">Compose</span></span><br><span data-ttu-id="1ad0d-142">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-142">Read</span></span> | [<span data-ttu-id="1ad0d-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="1ad0d-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="1ad0d-144">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-145">平台</span><span class="sxs-lookup"><span data-stu-id="1ad0d-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="1ad0d-146">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-146">Compose</span></span><br><span data-ttu-id="1ad0d-147">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-147">Read</span></span> | [<span data-ttu-id="1ad0d-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1ad0d-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="1ad0d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-150">满足</span><span class="sxs-lookup"><span data-stu-id="1ad0d-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="1ad0d-151">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-151">Compose</span></span><br><span data-ttu-id="1ad0d-152">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-152">Read</span></span> | [<span data-ttu-id="1ad0d-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1ad0d-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="1ad0d-154">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="1ad0d-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="1ad0d-156">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-156">Compose</span></span><br><span data-ttu-id="1ad0d-157">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-157">Read</span></span> | [<span data-ttu-id="1ad0d-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1ad0d-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="1ad0d-159">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ad0d-160">ui</span><span class="sxs-lookup"><span data-stu-id="1ad0d-160">ui</span></span>](#ui-ui) | <span data-ttu-id="1ad0d-161">撰写</span><span class="sxs-lookup"><span data-stu-id="1ad0d-161">Compose</span></span><br><span data-ttu-id="1ad0d-162">读取</span><span class="sxs-lookup"><span data-stu-id="1ad0d-162">Read</span></span> | [<span data-ttu-id="1ad0d-163">UI</span><span class="sxs-lookup"><span data-stu-id="1ad0d-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="1ad0d-164">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="1ad0d-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="1ad0d-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="1ad0d-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="1ad0d-166">contentLanguage: String</span></span>

<span data-ttu-id="1ad0d-167">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="1ad0d-168">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-169">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-169">Type</span></span>

*   <span data-ttu-id="1ad0d-170">String</span><span class="sxs-lookup"><span data-stu-id="1ad0d-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ad0d-171">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-171">Requirements</span></span>

|<span data-ttu-id="1ad0d-172">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-172">Requirement</span></span>| <span data-ttu-id="1ad0d-173">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-175">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-175">1.1</span></span>|
|[<span data-ttu-id="1ad0d-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ad0d-178">示例</span><span class="sxs-lookup"><span data-stu-id="1ad0d-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="1ad0d-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1ad0d-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="1ad0d-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-181">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-181">Type</span></span>

*   [<span data-ttu-id="1ad0d-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1ad0d-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="1ad0d-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ad0d-183">Requirements</span></span>

|<span data-ttu-id="1ad0d-184">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-184">Requirement</span></span>| <span data-ttu-id="1ad0d-185">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-187">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-187">1.1</span></span>|
|[<span data-ttu-id="1ad0d-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ad0d-190">示例</span><span class="sxs-lookup"><span data-stu-id="1ad0d-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="1ad0d-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="1ad0d-191">displayLanguage: String</span></span>

<span data-ttu-id="1ad0d-192">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="1ad0d-193">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-194">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-194">Type</span></span>

*   <span data-ttu-id="1ad0d-195">String</span><span class="sxs-lookup"><span data-stu-id="1ad0d-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ad0d-196">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-196">Requirements</span></span>

|<span data-ttu-id="1ad0d-197">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-197">Requirement</span></span>| <span data-ttu-id="1ad0d-198">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-200">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-200">1.1</span></span>|
|[<span data-ttu-id="1ad0d-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ad0d-203">示例</span><span class="sxs-lookup"><span data-stu-id="1ad0d-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="1ad0d-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="1ad0d-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="1ad0d-205">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-206">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-206">Type</span></span>

*   [<span data-ttu-id="1ad0d-207">HostType</span><span class="sxs-lookup"><span data-stu-id="1ad0d-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="1ad0d-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ad0d-208">Requirements</span></span>

|<span data-ttu-id="1ad0d-209">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-209">Requirement</span></span>| <span data-ttu-id="1ad0d-210">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-212">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-212">1.1</span></span>|
|[<span data-ttu-id="1ad0d-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-214">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ad0d-215">示例</span><span class="sxs-lookup"><span data-stu-id="1ad0d-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="1ad0d-216">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="1ad0d-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="1ad0d-217">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-218">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-218">Type</span></span>

*   [<span data-ttu-id="1ad0d-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1ad0d-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="1ad0d-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ad0d-220">Requirements</span></span>

|<span data-ttu-id="1ad0d-221">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-221">Requirement</span></span>| <span data-ttu-id="1ad0d-222">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-224">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-224">1.1</span></span>|
|[<span data-ttu-id="1ad0d-225">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-226">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ad0d-227">示例</span><span class="sxs-lookup"><span data-stu-id="1ad0d-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="1ad0d-228">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="1ad0d-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="1ad0d-229">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-230">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-230">Type</span></span>

*   [<span data-ttu-id="1ad0d-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1ad0d-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="1ad0d-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ad0d-232">Requirements</span></span>

|<span data-ttu-id="1ad0d-233">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-233">Requirement</span></span>| <span data-ttu-id="1ad0d-234">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-236">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-236">1.1</span></span>|
|[<span data-ttu-id="1ad0d-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-238">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ad0d-239">示例</span><span class="sxs-lookup"><span data-stu-id="1ad0d-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="1ad0d-240">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="1ad0d-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="1ad0d-241">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1ad0d-242">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-243">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-243">Type</span></span>

*   [<span data-ttu-id="1ad0d-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1ad0d-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1ad0d-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ad0d-245">Requirements</span></span>

|<span data-ttu-id="1ad0d-246">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-246">Requirement</span></span>| <span data-ttu-id="1ad0d-247">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-249">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-249">1.1</span></span>|
|[<span data-ttu-id="1ad0d-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1ad0d-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="1ad0d-251">受限</span><span class="sxs-lookup"><span data-stu-id="1ad0d-251">Restricted</span></span>|
|[<span data-ttu-id="1ad0d-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="1ad0d-254">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="1ad0d-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="1ad0d-255">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="1ad0d-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1ad0d-256">类型</span><span class="sxs-lookup"><span data-stu-id="1ad0d-256">Type</span></span>

*   [<span data-ttu-id="1ad0d-257">UI</span><span class="sxs-lookup"><span data-stu-id="1ad0d-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="1ad0d-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="1ad0d-258">Requirements</span></span>

|<span data-ttu-id="1ad0d-259">要求</span><span class="sxs-lookup"><span data-stu-id="1ad0d-259">Requirement</span></span>| <span data-ttu-id="1ad0d-260">值</span><span class="sxs-lookup"><span data-stu-id="1ad0d-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ad0d-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1ad0d-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ad0d-262">1.1</span><span class="sxs-lookup"><span data-stu-id="1ad0d-262">1.1</span></span>|
|[<span data-ttu-id="1ad0d-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1ad0d-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ad0d-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1ad0d-264">Compose or Read</span></span>|
