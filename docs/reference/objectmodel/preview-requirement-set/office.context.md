---
title: Office. context-预览要求集
description: Outlook 外接程序 API 中的 Outlook 上下文对象的对象模型（邮箱 API Preview 版本）。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 409f0a5b46eba667f79228f45081c160c3c3ce7f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717801"
---
# <a name="context"></a><span data-ttu-id="baca0-103">context</span><span class="sxs-lookup"><span data-stu-id="baca0-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="baca0-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="baca0-104">[Office](office.md).context</span></span>

<span data-ttu-id="baca0-105">在所有 Office 应用中，上下文提供外接程序使用的共享接口。</span><span class="sxs-lookup"><span data-stu-id="baca0-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="baca0-106">此列表仅记录 Outlook 外接程序使用的那些接口。有关 "context" 命名空间的完整列表，请参阅[通用 API 中的 "office. context reference](/javascript/api/office/office.context?view=outlook-js-preview)"。</span><span class="sxs-lookup"><span data-stu-id="baca0-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="baca0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-107">Requirements</span></span>

|<span data-ttu-id="baca0-108">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-108">Requirement</span></span>| <span data-ttu-id="baca0-109">值</span><span class="sxs-lookup"><span data-stu-id="baca0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-111">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-111">1.1</span></span>|
|[<span data-ttu-id="baca0-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="baca0-114">属性</span><span class="sxs-lookup"><span data-stu-id="baca0-114">Properties</span></span>

| <span data-ttu-id="baca0-115">属性</span><span class="sxs-lookup"><span data-stu-id="baca0-115">Property</span></span> | <span data-ttu-id="baca0-116">型号</span><span class="sxs-lookup"><span data-stu-id="baca0-116">Modes</span></span> | <span data-ttu-id="baca0-117">返回类型</span><span class="sxs-lookup"><span data-stu-id="baca0-117">Return type</span></span> | <span data-ttu-id="baca0-118">最低</span><span class="sxs-lookup"><span data-stu-id="baca0-118">Minimum</span></span><br><span data-ttu-id="baca0-119">要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="baca0-120">认证</span><span class="sxs-lookup"><span data-stu-id="baca0-120">auth</span></span>](#auth-auth) | <span data-ttu-id="baca0-121">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-121">Compose</span></span><br><span data-ttu-id="baca0-122">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-122">Read</span></span> | [<span data-ttu-id="baca0-123">Auth</span><span class="sxs-lookup"><span data-stu-id="baca0-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="baca0-124">预览</span><span class="sxs-lookup"><span data-stu-id="baca0-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="baca0-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="baca0-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="baca0-126">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-126">Compose</span></span><br><span data-ttu-id="baca0-127">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-127">Read</span></span> | <span data-ttu-id="baca0-128">String</span><span class="sxs-lookup"><span data-stu-id="baca0-128">String</span></span> | [<span data-ttu-id="baca0-129">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-130">过程</span><span class="sxs-lookup"><span data-stu-id="baca0-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="baca0-131">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-131">Compose</span></span><br><span data-ttu-id="baca0-132">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-132">Read</span></span> | [<span data-ttu-id="baca0-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="baca0-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="baca0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="baca0-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="baca0-136">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-136">Compose</span></span><br><span data-ttu-id="baca0-137">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-137">Read</span></span> | <span data-ttu-id="baca0-138">String</span><span class="sxs-lookup"><span data-stu-id="baca0-138">String</span></span> | [<span data-ttu-id="baca0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-140">host</span><span class="sxs-lookup"><span data-stu-id="baca0-140">host</span></span>](#host-hosttype) | <span data-ttu-id="baca0-141">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-141">Compose</span></span><br><span data-ttu-id="baca0-142">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-142">Read</span></span> | [<span data-ttu-id="baca0-143">HostType</span><span class="sxs-lookup"><span data-stu-id="baca0-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="baca0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="baca0-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="baca0-146">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-146">Compose</span></span><br><span data-ttu-id="baca0-147">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-147">Read</span></span> | [<span data-ttu-id="baca0-148">邮箱</span><span class="sxs-lookup"><span data-stu-id="baca0-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="baca0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="baca0-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="baca0-151">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-151">Compose</span></span><br><span data-ttu-id="baca0-152">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-152">Read</span></span> | [<span data-ttu-id="baca0-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="baca0-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="baca0-154">预览</span><span class="sxs-lookup"><span data-stu-id="baca0-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="baca0-155">平台</span><span class="sxs-lookup"><span data-stu-id="baca0-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="baca0-156">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-156">Compose</span></span><br><span data-ttu-id="baca0-157">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-157">Read</span></span> | [<span data-ttu-id="baca0-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="baca0-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="baca0-159">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-160">满足</span><span class="sxs-lookup"><span data-stu-id="baca0-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="baca0-161">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-161">Compose</span></span><br><span data-ttu-id="baca0-162">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-162">Read</span></span> | [<span data-ttu-id="baca0-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="baca0-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="baca0-164">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="baca0-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="baca0-166">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-166">Compose</span></span><br><span data-ttu-id="baca0-167">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-167">Read</span></span> | [<span data-ttu-id="baca0-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="baca0-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="baca0-169">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="baca0-170">ui</span><span class="sxs-lookup"><span data-stu-id="baca0-170">ui</span></span>](#ui-ui) | <span data-ttu-id="baca0-171">撰写</span><span class="sxs-lookup"><span data-stu-id="baca0-171">Compose</span></span><br><span data-ttu-id="baca0-172">读取</span><span class="sxs-lookup"><span data-stu-id="baca0-172">Read</span></span> | [<span data-ttu-id="baca0-173">UI</span><span class="sxs-lookup"><span data-stu-id="baca0-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="baca0-174">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="baca0-175">属性详细信息</span><span class="sxs-lookup"><span data-stu-id="baca0-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="baca0-176">auth： [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="baca0-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="baca0-177">通过提供一种方法来支持[单一登录（SSO）](../../../outlook/authenticate-a-user-with-an-sso-token.md) ，使 Office 主机能够获取加载项的 web 应用程序的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="baca0-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="baca0-178">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="baca0-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-179">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-179">Type</span></span>

*   [<span data-ttu-id="baca0-180">Auth</span><span class="sxs-lookup"><span data-stu-id="baca0-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="baca0-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-181">Requirements</span></span>

|<span data-ttu-id="baca0-182">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-182">Requirement</span></span>| <span data-ttu-id="baca0-183">值</span><span class="sxs-lookup"><span data-stu-id="baca0-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-184">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-185">预览</span><span class="sxs-lookup"><span data-stu-id="baca0-185">Preview</span></span>|
|[<span data-ttu-id="baca0-186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-187">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-188">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-188">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="baca0-189">contentLanguage： String</span><span class="sxs-lookup"><span data-stu-id="baca0-189">contentLanguage: String</span></span>

<span data-ttu-id="baca0-190">获取用户指定的用于编辑项的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="baca0-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="baca0-191">此`contentLanguage`值反映了在 Office 主机应用程序中使用**File > Options > 语言**指定的当前**编辑语言**设置。</span><span class="sxs-lookup"><span data-stu-id="baca0-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-192">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-192">Type</span></span>

*   <span data-ttu-id="baca0-193">String</span><span class="sxs-lookup"><span data-stu-id="baca0-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="baca0-194">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-194">Requirements</span></span>

|<span data-ttu-id="baca0-195">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-195">Requirement</span></span>| <span data-ttu-id="baca0-196">值</span><span class="sxs-lookup"><span data-stu-id="baca0-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-198">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-198">1.1</span></span>|
|[<span data-ttu-id="baca0-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-201">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="baca0-202">诊断： [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="baca0-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="baca0-203">获取有关加载项在其中运行的环境的信息。</span><span class="sxs-lookup"><span data-stu-id="baca0-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-204">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-204">Type</span></span>

*   [<span data-ttu-id="baca0-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="baca0-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="baca0-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-206">Requirements</span></span>

|<span data-ttu-id="baca0-207">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-207">Requirement</span></span>| <span data-ttu-id="baca0-208">值</span><span class="sxs-lookup"><span data-stu-id="baca0-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-210">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-210">1.1</span></span>|
|[<span data-ttu-id="baca0-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-212">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-213">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="baca0-214">displayLanguage： String</span><span class="sxs-lookup"><span data-stu-id="baca0-214">displayLanguage: String</span></span>

<span data-ttu-id="baca0-215">获取用户针对 Office 主机应用程序的 UI 指定的 RFC 1766 语言标记格式的区域设置（语言）。</span><span class="sxs-lookup"><span data-stu-id="baca0-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="baca0-216">`displayLanguage` 值反映在 Office 主机应用程序中通过“**文件 > 选项 > 语言**”指定的当前“**显示语言**”设置。</span><span class="sxs-lookup"><span data-stu-id="baca0-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-217">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-217">Type</span></span>

*   <span data-ttu-id="baca0-218">String</span><span class="sxs-lookup"><span data-stu-id="baca0-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="baca0-219">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-219">Requirements</span></span>

|<span data-ttu-id="baca0-220">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-220">Requirement</span></span>| <span data-ttu-id="baca0-221">值</span><span class="sxs-lookup"><span data-stu-id="baca0-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-223">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-223">1.1</span></span>|
|[<span data-ttu-id="baca0-224">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-225">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-226">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="baca0-227">主机： [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="baca0-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="baca0-228">获取运行外接程序的 Office 应用程序主机。</span><span class="sxs-lookup"><span data-stu-id="baca0-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-229">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-229">Type</span></span>

*   [<span data-ttu-id="baca0-230">HostType</span><span class="sxs-lookup"><span data-stu-id="baca0-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="baca0-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-231">Requirements</span></span>

|<span data-ttu-id="baca0-232">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-232">Requirement</span></span>| <span data-ttu-id="baca0-233">值</span><span class="sxs-lookup"><span data-stu-id="baca0-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-234">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-235">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-235">1.1</span></span>|
|[<span data-ttu-id="baca0-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-238">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="baca0-239">officeTheme： [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="baca0-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="baca0-240">提供了访问 Office 主题颜色的属性。</span><span class="sxs-lookup"><span data-stu-id="baca0-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="baca0-241">此成员仅在 Windows 中的 Outlook 中受支持。</span><span class="sxs-lookup"><span data-stu-id="baca0-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="baca0-242">使用 Office 主题颜色，您可以将加载项的配色方案与用户选择的当前 Office 主题进行协调，以供用户使用**office > Office 帐户 > Office 主题 UI**，该用户在所有 Office 主机应用程序中应用。</span><span class="sxs-lookup"><span data-stu-id="baca0-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="baca0-243">使用 Office 主题颜色适用于邮件和任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="baca0-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-244">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-244">Type</span></span>

*   [<span data-ttu-id="baca0-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="baca0-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="baca0-246">属性：</span><span class="sxs-lookup"><span data-stu-id="baca0-246">Properties:</span></span>

|<span data-ttu-id="baca0-247">姓名</span><span class="sxs-lookup"><span data-stu-id="baca0-247">Name</span></span>| <span data-ttu-id="baca0-248">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-248">Type</span></span>| <span data-ttu-id="baca0-249">说明</span><span class="sxs-lookup"><span data-stu-id="baca0-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="baca0-250">String</span><span class="sxs-lookup"><span data-stu-id="baca0-250">String</span></span>|<span data-ttu-id="baca0-251">获取十六进制三原色形式的 Office 主题正文背景色。</span><span class="sxs-lookup"><span data-stu-id="baca0-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="baca0-252">String</span><span class="sxs-lookup"><span data-stu-id="baca0-252">String</span></span>|<span data-ttu-id="baca0-253">获取十六进制三原色形式的 Office 主题正文前景色。</span><span class="sxs-lookup"><span data-stu-id="baca0-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="baca0-254">String</span><span class="sxs-lookup"><span data-stu-id="baca0-254">String</span></span>|<span data-ttu-id="baca0-255">获取十六进制三原色形式的 Office 主题控制背景色。</span><span class="sxs-lookup"><span data-stu-id="baca0-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="baca0-256">字符串</span><span class="sxs-lookup"><span data-stu-id="baca0-256">String</span></span>|<span data-ttu-id="baca0-257">获取十六进制三原色形式的 Office 主题正文控制颜色。</span><span class="sxs-lookup"><span data-stu-id="baca0-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="baca0-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-258">Requirements</span></span>

|<span data-ttu-id="baca0-259">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-259">Requirement</span></span>| <span data-ttu-id="baca0-260">值</span><span class="sxs-lookup"><span data-stu-id="baca0-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-262">预览</span><span class="sxs-lookup"><span data-stu-id="baca0-262">Preview</span></span>|
|[<span data-ttu-id="baca0-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-265">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-265">Example</span></span>

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

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="baca0-266">platform： [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="baca0-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="baca0-267">提供在其上运行外接的平台。</span><span class="sxs-lookup"><span data-stu-id="baca0-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-268">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-268">Type</span></span>

*   [<span data-ttu-id="baca0-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="baca0-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="baca0-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-270">Requirements</span></span>

|<span data-ttu-id="baca0-271">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-271">Requirement</span></span>| <span data-ttu-id="baca0-272">值</span><span class="sxs-lookup"><span data-stu-id="baca0-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-273">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-274">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-274">1.1</span></span>|
|[<span data-ttu-id="baca0-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-277">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="baca0-278">要求： [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="baca0-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="baca0-279">提供用于确定当前主机和平台上支持的要求集的方法。</span><span class="sxs-lookup"><span data-stu-id="baca0-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-280">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-280">Type</span></span>

*   [<span data-ttu-id="baca0-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="baca0-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="baca0-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-282">Requirements</span></span>

|<span data-ttu-id="baca0-283">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-283">Requirement</span></span>| <span data-ttu-id="baca0-284">值</span><span class="sxs-lookup"><span data-stu-id="baca0-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-285">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-286">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-286">1.1</span></span>|
|[<span data-ttu-id="baca0-287">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-288">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="baca0-289">示例</span><span class="sxs-lookup"><span data-stu-id="baca0-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="baca0-290">roamingSettings： [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="baca0-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="baca0-291">获取一个对象，它表示保存到用户邮箱的邮件外接程序的自定义设置或状态。</span><span class="sxs-lookup"><span data-stu-id="baca0-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="baca0-292">`RoamingSettings` 对象允许您存储和访问用户邮箱中存储的邮件外接程序的数据，以便从用于访问该邮箱的任何主机客户端应用程序中运行该外接程序时，该外接程序可以使用该数据。</span><span class="sxs-lookup"><span data-stu-id="baca0-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-293">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-293">Type</span></span>

*   [<span data-ttu-id="baca0-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="baca0-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="baca0-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-295">Requirements</span></span>

|<span data-ttu-id="baca0-296">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-296">Requirement</span></span>| <span data-ttu-id="baca0-297">值</span><span class="sxs-lookup"><span data-stu-id="baca0-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-299">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-299">1.1</span></span>|
|[<span data-ttu-id="baca0-300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="baca0-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="baca0-301">受限</span><span class="sxs-lookup"><span data-stu-id="baca0-301">Restricted</span></span>|
|[<span data-ttu-id="baca0-302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-303">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="baca0-304">ui： [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="baca0-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="baca0-305">提供可用于在 Office 外接程序中创建和操作 UI 组件（如对话框）的对象和方法。</span><span class="sxs-lookup"><span data-stu-id="baca0-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="baca0-306">类型</span><span class="sxs-lookup"><span data-stu-id="baca0-306">Type</span></span>

*   [<span data-ttu-id="baca0-307">UI</span><span class="sxs-lookup"><span data-stu-id="baca0-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="baca0-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="baca0-308">Requirements</span></span>

|<span data-ttu-id="baca0-309">要求</span><span class="sxs-lookup"><span data-stu-id="baca0-309">Requirement</span></span>| <span data-ttu-id="baca0-310">值</span><span class="sxs-lookup"><span data-stu-id="baca0-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="baca0-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="baca0-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="baca0-312">1.1</span><span class="sxs-lookup"><span data-stu-id="baca0-312">1.1</span></span>|
|[<span data-ttu-id="baca0-313">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="baca0-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="baca0-314">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="baca0-314">Compose or Read</span></span>|
