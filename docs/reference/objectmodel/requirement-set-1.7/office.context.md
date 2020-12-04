---
title: Office。上下文要求集1。7
description: 使用邮箱 API 要求集1.7 的 Outlook 外接程序可用的 Office 对象成员。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 4a1ca6b4975ffba2c2bd400267fbe7db63f88244
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570728"
---
# <a name="context-mailbox-requirement-set-17"></a><span data-ttu-id="390dd-103"> (邮箱要求集1.7 的上下文) </span><span class="sxs-lookup"><span data-stu-id="390dd-103">context (Mailbox requirement set 1.7)</span></span>

### <a name="officecontext"></a><span data-ttu-id="390dd-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="390dd-104">[Office](office.md).context</span></span>

<span data-ttu-id="390dd-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="390dd-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="390dd-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅 [通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true)"。</span><span class="sxs-lookup"><span data-stu-id="390dd-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="390dd-107">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-107">Requirements</span></span>

|<span data-ttu-id="390dd-108">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-108">Requirement</span></span>| <span data-ttu-id="390dd-109">值</span><span class="sxs-lookup"><span data-stu-id="390dd-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-111">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-111">1.1</span></span>|
|[<span data-ttu-id="390dd-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="390dd-114">属性</span><span class="sxs-lookup"><span data-stu-id="390dd-114">Properties</span></span>

| <span data-ttu-id="390dd-115">属性</span><span class="sxs-lookup"><span data-stu-id="390dd-115">Property</span></span> | <span data-ttu-id="390dd-116">型号</span><span class="sxs-lookup"><span data-stu-id="390dd-116">Modes</span></span> | <span data-ttu-id="390dd-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="390dd-117">Return type</span></span> | <span data-ttu-id="390dd-118">最小值</span><span class="sxs-lookup"><span data-stu-id="390dd-118">Minimum</span></span><br><span data-ttu-id="390dd-119">要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="390dd-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="390dd-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="390dd-121">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-121">Compose</span></span><br><span data-ttu-id="390dd-122">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-122">Read</span></span> | <span data-ttu-id="390dd-123">String</span><span class="sxs-lookup"><span data-stu-id="390dd-123">String</span></span> | [<span data-ttu-id="390dd-124">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="390dd-125">过程</span><span class="sxs-lookup"><span data-stu-id="390dd-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="390dd-126">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-126">Compose</span></span><br><span data-ttu-id="390dd-127">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-127">Read</span></span> | [<span data-ttu-id="390dd-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="390dd-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="390dd-129">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="390dd-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="390dd-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="390dd-131">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-131">Compose</span></span><br><span data-ttu-id="390dd-132">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-132">Read</span></span> | <span data-ttu-id="390dd-133">String</span><span class="sxs-lookup"><span data-stu-id="390dd-133">String</span></span> | [<span data-ttu-id="390dd-134">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="390dd-135">host</span><span class="sxs-lookup"><span data-stu-id="390dd-135">host</span></span>](#host-hosttype) | <span data-ttu-id="390dd-136">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-136">Compose</span></span><br><span data-ttu-id="390dd-137">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-137">Read</span></span> | [<span data-ttu-id="390dd-138">HostType</span><span class="sxs-lookup"><span data-stu-id="390dd-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="390dd-139">1.5</span><span class="sxs-lookup"><span data-stu-id="390dd-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="390dd-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="390dd-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="390dd-141">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-141">Compose</span></span><br><span data-ttu-id="390dd-142">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-142">Read</span></span> | [<span data-ttu-id="390dd-143">邮箱</span><span class="sxs-lookup"><span data-stu-id="390dd-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="390dd-144">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="390dd-145">平台</span><span class="sxs-lookup"><span data-stu-id="390dd-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="390dd-146">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-146">Compose</span></span><br><span data-ttu-id="390dd-147">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-147">Read</span></span> | [<span data-ttu-id="390dd-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="390dd-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="390dd-149">1.5</span><span class="sxs-lookup"><span data-stu-id="390dd-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="390dd-150">满足</span><span class="sxs-lookup"><span data-stu-id="390dd-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="390dd-151">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-151">Compose</span></span><br><span data-ttu-id="390dd-152">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-152">Read</span></span> | [<span data-ttu-id="390dd-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="390dd-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="390dd-154">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="390dd-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="390dd-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="390dd-156">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-156">Compose</span></span><br><span data-ttu-id="390dd-157">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-157">Read</span></span> | [<span data-ttu-id="390dd-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="390dd-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="390dd-159">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="390dd-160">ui</span><span class="sxs-lookup"><span data-stu-id="390dd-160">ui</span></span>](#ui-ui) | <span data-ttu-id="390dd-161">撰写</span><span class="sxs-lookup"><span data-stu-id="390dd-161">Compose</span></span><br><span data-ttu-id="390dd-162">阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-162">Read</span></span> | [<span data-ttu-id="390dd-163">UI</span><span class="sxs-lookup"><span data-stu-id="390dd-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="390dd-164">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="390dd-165">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="390dd-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="390dd-166">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="390dd-166">contentLanguage: String</span></span>

<span data-ttu-id="390dd-167">获取用户指定的用于编辑项目的区域设置 (语言) 。</span><span class="sxs-lookup"><span data-stu-id="390dd-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="390dd-168">此 `contentLanguage` 值反映了使用 Office 客户端应用程序中的 "**文件 > 选项" > 语言** 指定的当前 **编辑语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="390dd-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-169">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-169">Type</span></span>

*   <span data-ttu-id="390dd-170">String</span><span class="sxs-lookup"><span data-stu-id="390dd-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="390dd-171">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-171">Requirements</span></span>

|<span data-ttu-id="390dd-172">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-172">Requirement</span></span>| <span data-ttu-id="390dd-173">值</span><span class="sxs-lookup"><span data-stu-id="390dd-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-174">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-175">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-175">1.1</span></span>|
|[<span data-ttu-id="390dd-176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-177">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="390dd-178">示例</span><span class="sxs-lookup"><span data-stu-id="390dd-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="390dd-179">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="390dd-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="390dd-180">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="390dd-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-181">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-181">Type</span></span>

*   [<span data-ttu-id="390dd-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="390dd-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="390dd-183">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-183">Requirements</span></span>

|<span data-ttu-id="390dd-184">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-184">Requirement</span></span>| <span data-ttu-id="390dd-185">值</span><span class="sxs-lookup"><span data-stu-id="390dd-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-186">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-187">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-187">1.1</span></span>|
|[<span data-ttu-id="390dd-188">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-189">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="390dd-190">示例</span><span class="sxs-lookup"><span data-stu-id="390dd-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="390dd-191">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="390dd-191">displayLanguage: String</span></span>

<span data-ttu-id="390dd-192">获取用户为 Office 客户端应用程序的 UI 指定的 RFC 1766 语言标记格式中 (语言) 的区域设置。</span><span class="sxs-lookup"><span data-stu-id="390dd-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="390dd-193">此 `displayLanguage` 值反映了使用 Office 客户端应用程序中的 **File > Options > 语言** 指定的当前 **显示语言** 设置。</span><span class="sxs-lookup"><span data-stu-id="390dd-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-194">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-194">Type</span></span>

*   <span data-ttu-id="390dd-195">String</span><span class="sxs-lookup"><span data-stu-id="390dd-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="390dd-196">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-196">Requirements</span></span>

|<span data-ttu-id="390dd-197">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-197">Requirement</span></span>| <span data-ttu-id="390dd-198">值</span><span class="sxs-lookup"><span data-stu-id="390dd-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-199">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-200">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-200">1.1</span></span>|
|[<span data-ttu-id="390dd-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="390dd-203">示例</span><span class="sxs-lookup"><span data-stu-id="390dd-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="390dd-204">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="390dd-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="390dd-205">获取承载外接程序的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="390dd-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="390dd-206">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取主机。</span><span class="sxs-lookup"><span data-stu-id="390dd-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-207">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-207">Type</span></span>

*   [<span data-ttu-id="390dd-208">HostType</span><span class="sxs-lookup"><span data-stu-id="390dd-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="390dd-209">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-209">Requirements</span></span>

|<span data-ttu-id="390dd-210">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-210">Requirement</span></span>| <span data-ttu-id="390dd-211">值</span><span class="sxs-lookup"><span data-stu-id="390dd-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-212">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-213">1.5</span><span class="sxs-lookup"><span data-stu-id="390dd-213">1.5</span></span>|
|[<span data-ttu-id="390dd-214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-215">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="390dd-216">示例</span><span class="sxs-lookup"><span data-stu-id="390dd-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="390dd-217">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="390dd-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="390dd-218">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="390dd-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="390dd-219">或者，也可以使用 " [context.subname](#diagnostics-contextinformation) " 属性获取平台。</span><span class="sxs-lookup"><span data-stu-id="390dd-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-220">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-220">Type</span></span>

*   [<span data-ttu-id="390dd-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="390dd-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="390dd-222">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-222">Requirements</span></span>

|<span data-ttu-id="390dd-223">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-223">Requirement</span></span>| <span data-ttu-id="390dd-224">值</span><span class="sxs-lookup"><span data-stu-id="390dd-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-226">1.5</span><span class="sxs-lookup"><span data-stu-id="390dd-226">1.5</span></span>|
|[<span data-ttu-id="390dd-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-228">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="390dd-229">示例</span><span class="sxs-lookup"><span data-stu-id="390dd-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="390dd-230">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="390dd-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="390dd-231">提供用于确定当前应用程序和平台支持哪些要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="390dd-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-232">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-232">Type</span></span>

*   [<span data-ttu-id="390dd-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="390dd-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="390dd-234">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-234">Requirements</span></span>

|<span data-ttu-id="390dd-235">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-235">Requirement</span></span>| <span data-ttu-id="390dd-236">值</span><span class="sxs-lookup"><span data-stu-id="390dd-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-238">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-238">1.1</span></span>|
|[<span data-ttu-id="390dd-239">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-240">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="390dd-241">示例</span><span class="sxs-lookup"><span data-stu-id="390dd-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="390dd-242">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="390dd-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="390dd-243">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="390dd-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="390dd-244">该 `RoamingSettings` 对象使您可以存储和访问存储在用户邮箱中的邮件外接程序的数据，以便该外接程序从用于访问该邮箱的任何 Outlook 客户端运行时都可使用该加载项。</span><span class="sxs-lookup"><span data-stu-id="390dd-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-245">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-245">Type</span></span>

*   [<span data-ttu-id="390dd-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="390dd-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="390dd-247">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-247">Requirements</span></span>

|<span data-ttu-id="390dd-248">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-248">Requirement</span></span>| <span data-ttu-id="390dd-249">值</span><span class="sxs-lookup"><span data-stu-id="390dd-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-250">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-251">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-251">1.1</span></span>|
|[<span data-ttu-id="390dd-252">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="390dd-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="390dd-253">受限</span><span class="sxs-lookup"><span data-stu-id="390dd-253">Restricted</span></span>|
|[<span data-ttu-id="390dd-254">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-255">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="390dd-256">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="390dd-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="390dd-257">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="390dd-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="390dd-258">类型</span><span class="sxs-lookup"><span data-stu-id="390dd-258">Type</span></span>

*   [<span data-ttu-id="390dd-259">UI</span><span class="sxs-lookup"><span data-stu-id="390dd-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="390dd-260">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-260">Requirements</span></span>

|<span data-ttu-id="390dd-261">要求</span><span class="sxs-lookup"><span data-stu-id="390dd-261">Requirement</span></span>| <span data-ttu-id="390dd-262">值</span><span class="sxs-lookup"><span data-stu-id="390dd-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="390dd-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="390dd-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="390dd-264">1.1</span><span class="sxs-lookup"><span data-stu-id="390dd-264">1.1</span></span>|
|[<span data-ttu-id="390dd-265">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="390dd-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="390dd-266">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="390dd-266">Compose or Read</span></span>|
