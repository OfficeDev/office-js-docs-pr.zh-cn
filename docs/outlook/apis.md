---
title: Outlook 加载项 API
description: 了解如何引用 Outlook 加载项 API 并声明 Outlook 加载项中的权限。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 2abe365f1606789b1c6ac113b133019055767b28
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166040"
---
# <a name="outlook-add-in-apis"></a><span data-ttu-id="bb3cd-103">Outlook 外接程序 API</span><span class="sxs-lookup"><span data-stu-id="bb3cd-103">Outlook add-in APIs</span></span>

<span data-ttu-id="bb3cd-104">要将 API 用于您的 Outlook 外接程序，您必须指定 Office.js 库的位置、要求集、架构和权限。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-104">To use APIs in your Outlook add-in, you must specify the location of the Office.js library, the requirement set, the schema, and the permissions.</span></span>

## <a name="officejs-library"></a><span data-ttu-id="bb3cd-105">Office.js 库</span><span class="sxs-lookup"><span data-stu-id="bb3cd-105">Office.js library</span></span>

<span data-ttu-id="bb3cd-106">若要与 Outlook 加载项 API 进行交互，需要在 Office.js 中使用 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-106">To interact with the Outlook add-in API, you need to use the JavaScript APIs in Office.js.</span></span> <span data-ttu-id="bb3cd-107">库的 CDN 为 `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-107">The CDN for the library is `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="bb3cd-108">提交到 AppSource 的加载项必须按此 CDN 引用 Office.js，它们不能使用本地引用。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-108">Add-ins submitted to AppSource must reference Office.js by this CDN; they can't use a local reference.</span></span>

<span data-ttu-id="bb3cd-109">在实现加载项 UI 的网页（.html、.aspx 或 .php 文件）的 `<head>` 标记的 `<script>` 标记中引用 CDN。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-109">Reference the CDN in a `<script>` tag in the `<head>` tag of the web page (.html, .aspx, or .php file) that implements the UI of your add-in.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```
<span data-ttu-id="bb3cd-p102">添加 API 时，Office.js 的 URL 将保持不变。仅当我们打破现有的 API 行为时，才会更改 URL 中的版本。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-p102">As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bb3cd-112">开发任何 Office 主机应用程序的加载项时，从页面的 `<head>` 区域内引用适用于 Office 的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-112">When developing an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page.</span></span> <span data-ttu-id="bb3cd-113">这样可确保 API 先于所有正文元素完全初始化。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-113">This ensures that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="bb3cd-114">Office 主机要求外接程序在激活 5 秒钟内进行初始化。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-114">Office hosts require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="bb3cd-115">超过此阈值会导致声明的外接程序无响应，并且会向用户显示错误消息。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-115">Crossing this threshold results in the add-in being declared unresponsive and an error message is displayed to the user.</span></span>

## <a name="requirement-sets"></a><span data-ttu-id="bb3cd-116">要求集</span><span class="sxs-lookup"><span data-stu-id="bb3cd-116">Requirement sets</span></span>

<span data-ttu-id="bb3cd-117">所有 Outlook API 都属于 `Mailbox` 要求集。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-117">All Outlook APIs belong to the `Mailbox` requirement set.</span></span> <span data-ttu-id="bb3cd-118">`Mailbox` 要求集具有多个版本，并且我们发布的每组新的 API 都属于更高版本的要求集。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-118">The `Mailbox` requirement set has versions, and each new set of APIs that are released belongs to a higher version of the set.</span></span> <span data-ttu-id="bb3cd-119">并非所有的 Outlook 客户端在发布时都将支持最新的 API 集，但如果 Outlook 客户端声明支持要求集，它将支持该要求集中的所有 API。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-119">Not all Outlook clients will support the newest set of APIs when they are released, but if an Outlook client declares support for a requirement set, it will support all the APIs in that requirement set.</span></span>

<span data-ttu-id="bb3cd-p105">若要控制外接程序在哪些 Outlook 客户端中显示，请在清单中指定最低要求集版本。例如，如果你指定要求集版本 1.3，则外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-p105">To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in will not show up in any Outlook client that doesn't support a minimum version of 1.3.</span></span>

<span data-ttu-id="bb3cd-p106">指定要求集不会将外接程序限定于该版本中的 API。如果外接程序指定要求集 v1.1，却在支持 v1.3 的 Outlook 客户端中运行，该外接程序仍可以使用 v1.3 API。要求集仅控制外接程序在哪些 Outlook 客户端中显示。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-p106">Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.</span></span>

<span data-ttu-id="bb3cd-125">要检查大于清单中所指定要求集的要求集中任何 API 的可用性，可以使用标准 JavaScrip：</span><span class="sxs-lookup"><span data-stu-id="bb3cd-125">To check the availability of any APIs from a requirement set greater than the one specified in the manifest, you can use standard JavaScript:</span></span>

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> <span data-ttu-id="bb3cd-126">对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-126">These checks are not needed for any APIs that are in the requirement set version specified in the manifest.</span></span>

<span data-ttu-id="bb3cd-127">指定支持你的方案的关键 API 集的最低要求集，如果缺少该要求集，加载项的功能将无法工作。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-127">Specify the minimum requirement set that supports the critical set of APIs for your scenario, without which features of your add-in won't work.</span></span> <span data-ttu-id="bb3cd-128">指定 `<Requirements>` 元素的清单中的要求集。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-128">You specify the requirement set in the manifest in the `<Requirements>` element.</span></span> <span data-ttu-id="bb3cd-129">有关更多信息，请参阅 [Outlook 加载项清单](manifests.md)和[了解 Outlook API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-129">For more information, see [Outlook add-in manifests](manifests.md) and [Understanding Outlook API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md).</span></span>

<span data-ttu-id="bb3cd-130">`<Methods>` 元素不适用于 Outlook 加载项，因此，无法声明对特定方法的支持。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-130">The `<Methods>` element doesn't apply to Outlook add-ins, so you can't declare support for specific methods.</span></span>

## <a name="permissions"></a><span data-ttu-id="bb3cd-131">权限</span><span class="sxs-lookup"><span data-stu-id="bb3cd-131">Permissions</span></span>

<span data-ttu-id="bb3cd-p108">外接程序需要相应的权限才能使用所需的 API。有四个级别的权限。有关详细信息，请参阅[了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-p108">Your add-in requires the appropriate permissions to use the APIs that it needs. There are four levels of permissions. For more details, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md).</span></span>

<br/>

|<span data-ttu-id="bb3cd-135">权限级别</span><span class="sxs-lookup"><span data-stu-id="bb3cd-135">Permission level</span></span>|<span data-ttu-id="bb3cd-136">说明</span><span class="sxs-lookup"><span data-stu-id="bb3cd-136">Description</span></span>|
|:-----|:-----|
| <span data-ttu-id="bb3cd-137">**受限**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-137">**Restricted**</span></span> | <span data-ttu-id="bb3cd-138">允许使用实体，但不允许使用正则表达式。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-138">Allows use of entities but not regular expressions.</span></span> |
| <span data-ttu-id="bb3cd-139">**读取项**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-139">**Read item**</span></span> | <span data-ttu-id="bb3cd-140">除了**受限**所允许的权限，它还允许：</span><span class="sxs-lookup"><span data-stu-id="bb3cd-140">In addition to what is allowed in **Restricted**, it allows:</span></span><ul><li><span data-ttu-id="bb3cd-141">正则表达式</span><span class="sxs-lookup"><span data-stu-id="bb3cd-141">regular expressions</span></span></li><li><span data-ttu-id="bb3cd-142">Outlook 外接程序 API 读取访问</span><span class="sxs-lookup"><span data-stu-id="bb3cd-142">Outlook add-in API read access</span></span></li><li><span data-ttu-id="bb3cd-143">获取项属性和回调令牌</span><span class="sxs-lookup"><span data-stu-id="bb3cd-143">getting the item properties and the callback token</span></span></li></ul> |
| <span data-ttu-id="bb3cd-144">**读/写**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-144">**Read/write**</span></span> | <span data-ttu-id="bb3cd-145">除了**读取项**所允许的权限，它还允许：</span><span class="sxs-lookup"><span data-stu-id="bb3cd-145">In addition to what is allowed in **Read item**, it allows:</span></span><ul><li><span data-ttu-id="bb3cd-146">Outlook 加载项 API 的完全访问权限，但不包括 `makeEwsRequestAsync`</span><span class="sxs-lookup"><span data-stu-id="bb3cd-146">full Outlook add-in API access except `makeEwsRequestAsync`</span></span></li><li><span data-ttu-id="bb3cd-147">设置项属性</span><span class="sxs-lookup"><span data-stu-id="bb3cd-147">setting the item properties</span></span></li></ul> |
| <span data-ttu-id="bb3cd-148">**读/写邮箱**</span><span class="sxs-lookup"><span data-stu-id="bb3cd-148">**Read/write mailbox**</span></span> | <span data-ttu-id="bb3cd-149">除了**读/写**所允许的权限，它还允许：</span><span class="sxs-lookup"><span data-stu-id="bb3cd-149">In addition to what is allowed in **Read/write**, it allows:</span></span><ul><li><span data-ttu-id="bb3cd-150">创建、读取、写入项和文件夹</span><span class="sxs-lookup"><span data-stu-id="bb3cd-150">creating, reading, writing items and folders</span></span></li><li><span data-ttu-id="bb3cd-151">发送项目</span><span class="sxs-lookup"><span data-stu-id="bb3cd-151">sending items</span></span></li><li><span data-ttu-id="bb3cd-152">调用 [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)</span><span class="sxs-lookup"><span data-stu-id="bb3cd-152">calling [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)</span></span></li></ul> |

<span data-ttu-id="bb3cd-153">一般情况下，应该指定加载项所需的最小权限。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-153">In general, you should specify the minimum permission needed for your add-in.</span></span> <span data-ttu-id="bb3cd-154">权限在清单的 `<Permissions>` 元素中声明。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-154">Permissions are declared in the `<Permissions>` element in the manifest.</span></span> <span data-ttu-id="bb3cd-155">有关更多信息，请参阅 [Outlook 加载项清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-155">For more information, see [Outlook add-in manifests](manifests.md).</span></span> <span data-ttu-id="bb3cd-156">有关安全问题的信息，请参阅[Office 外接程序的隐私和安全性](../concepts/privacy-and-security.md)。</span><span class="sxs-lookup"><span data-stu-id="bb3cd-156">For information about security issues, see [Privacy and security for Office Add-ins](../concepts/privacy-and-security.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="bb3cd-157">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bb3cd-157">See also</span></span>

- [<span data-ttu-id="bb3cd-158">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="bb3cd-158">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="bb3cd-159">了解 Outlook API 要求集</span><span class="sxs-lookup"><span data-stu-id="bb3cd-159">Understanding Outlook API requirement sets</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [<span data-ttu-id="bb3cd-160">Office 加载项的隐私和安全</span><span class="sxs-lookup"><span data-stu-id="bb3cd-160">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
