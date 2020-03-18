---
title: 在 Outlook 加载项中启用代理访问方案
description: 简要介绍了代理访问权限，并讨论了如何配置加载项支持。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 0941e4f0b5e1082b8a762acfa013d4e58be03469
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721014"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="256f0-103">在 Outlook 加载项中启用代理访问方案</span><span class="sxs-lookup"><span data-stu-id="256f0-103">Enable delegate access scenarios in an Outlook add-in</span></span>

<span data-ttu-id="256f0-104">邮箱所有者可以使用代理访问功能，以[允许其他人管理其邮件和日历](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)。</span><span class="sxs-lookup"><span data-stu-id="256f0-104">A mailbox owner can use the delegate access feature to [allow someone else to manage their mail and calendar](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="256f0-105">本文指定 Office JavaScript API 支持的代理权限，并介绍如何在 Outlook 外接程序中启用代理访问方案。</span><span class="sxs-lookup"><span data-stu-id="256f0-105">This article specifies which delegate permissions the Office JavaScript API supports and describes how to enable delegate access scenarios in your Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="256f0-106">代理访问当前在 Mac、Android 和 iOS 的 Outlook 中不可用。</span><span class="sxs-lookup"><span data-stu-id="256f0-106">Delegate access is not currently available in Outlook on Mac, Android, and iOS.</span></span> <span data-ttu-id="256f0-107">将来可提供此功能。</span><span class="sxs-lookup"><span data-stu-id="256f0-107">This functionality may be made available in the future.</span></span>
>
> <span data-ttu-id="256f0-108">对此功能的支持是在要求集1.8 中引入的。</span><span class="sxs-lookup"><span data-stu-id="256f0-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="256f0-109">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="256f0-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-permissions-for-delegate-access"></a><span data-ttu-id="256f0-110">代理访问支持的权限</span><span class="sxs-lookup"><span data-stu-id="256f0-110">Supported permissions for delegate access</span></span>

<span data-ttu-id="256f0-111">下表介绍了 Office JavaScript API 支持的代理权限。</span><span class="sxs-lookup"><span data-stu-id="256f0-111">The following table describes the delegate permissions that the Office JavaScript API supports.</span></span>

|<span data-ttu-id="256f0-112">权限</span><span class="sxs-lookup"><span data-stu-id="256f0-112">Permission</span></span>|<span data-ttu-id="256f0-113">值</span><span class="sxs-lookup"><span data-stu-id="256f0-113">Value</span></span>|<span data-ttu-id="256f0-114">说明</span><span class="sxs-lookup"><span data-stu-id="256f0-114">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="256f0-115">读取</span><span class="sxs-lookup"><span data-stu-id="256f0-115">Read</span></span>|<span data-ttu-id="256f0-116">1（000001）</span><span class="sxs-lookup"><span data-stu-id="256f0-116">1 (000001)</span></span>|<span data-ttu-id="256f0-117">可以读取项目。</span><span class="sxs-lookup"><span data-stu-id="256f0-117">Can read items.</span></span>|
|<span data-ttu-id="256f0-118">写入</span><span class="sxs-lookup"><span data-stu-id="256f0-118">Write</span></span>|<span data-ttu-id="256f0-119">2（000010）</span><span class="sxs-lookup"><span data-stu-id="256f0-119">2 (000010)</span></span>|<span data-ttu-id="256f0-120">可以创建项目。</span><span class="sxs-lookup"><span data-stu-id="256f0-120">Can create items.</span></span>|
|<span data-ttu-id="256f0-121">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="256f0-121">DeleteOwn</span></span>|<span data-ttu-id="256f0-122">4（000100）</span><span class="sxs-lookup"><span data-stu-id="256f0-122">4 (000100)</span></span>|<span data-ttu-id="256f0-123">只能删除其创建的项目。</span><span class="sxs-lookup"><span data-stu-id="256f0-123">Can delete only the items they created.</span></span>|
|<span data-ttu-id="256f0-124">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="256f0-124">DeleteAll</span></span>|<span data-ttu-id="256f0-125">8（001000）</span><span class="sxs-lookup"><span data-stu-id="256f0-125">8 (001000)</span></span>|<span data-ttu-id="256f0-126">可以删除任何项目。</span><span class="sxs-lookup"><span data-stu-id="256f0-126">Can delete any items.</span></span>|
|<span data-ttu-id="256f0-127">EditOwn</span><span class="sxs-lookup"><span data-stu-id="256f0-127">EditOwn</span></span>|<span data-ttu-id="256f0-128">16（010000）</span><span class="sxs-lookup"><span data-stu-id="256f0-128">16 (010000)</span></span>|<span data-ttu-id="256f0-129">只能编辑其创建的项目。</span><span class="sxs-lookup"><span data-stu-id="256f0-129">Can edit only the items they created.</span></span>|
|<span data-ttu-id="256f0-130">EditAll</span><span class="sxs-lookup"><span data-stu-id="256f0-130">EditAll</span></span>|<span data-ttu-id="256f0-131">32（100000）</span><span class="sxs-lookup"><span data-stu-id="256f0-131">32 (100000)</span></span>|<span data-ttu-id="256f0-132">可以编辑任何项目。</span><span class="sxs-lookup"><span data-stu-id="256f0-132">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="256f0-133">目前，API 支持获取现有的代理权限，但不支持设置委派权限。</span><span class="sxs-lookup"><span data-stu-id="256f0-133">Currently the API supports getting existing delegate permissions, but not setting delegate permissions.</span></span>

<span data-ttu-id="256f0-134">使用位掩码来实现[DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)对象，以指示代理的权限。</span><span class="sxs-lookup"><span data-stu-id="256f0-134">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the delegate's permissions.</span></span> <span data-ttu-id="256f0-135">位掩码中的每个位置都代表一个特定权限，如果将`1`其设置为，则代理具有相应的权限。</span><span class="sxs-lookup"><span data-stu-id="256f0-135">Each position in the bitmask represents a particular permission and if it's set to `1` then the delegate has the respective permission.</span></span> <span data-ttu-id="256f0-136">例如，如果右边的第二位是`1`，则代理具有 "**写入**" 权限。</span><span class="sxs-lookup"><span data-stu-id="256f0-136">For example, if the second bit from the right is `1`, then the delegate has **Write** permission.</span></span> <span data-ttu-id="256f0-137">您可以在本文后面的 "将[操作作为代理执行操作](#perform-an-operation-as-delegate)" 一节中查看有关如何检查特定权限的示例。</span><span class="sxs-lookup"><span data-stu-id="256f0-137">You can see an example of how to check for a specific permission in the [Perform an operation as delegate](#perform-an-operation-as-delegate) section later in this article.</span></span>

## <a name="sync-across-mailbox-clients"></a><span data-ttu-id="256f0-138">在邮箱客户端之间同步</span><span class="sxs-lookup"><span data-stu-id="256f0-138">Sync across mailbox clients</span></span>

<span data-ttu-id="256f0-139">代理对所有者邮箱的更新通常会在邮箱之间立即同步。</span><span class="sxs-lookup"><span data-stu-id="256f0-139">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="256f0-140">但是，如果外接程序使用 REST 或 EWS 操作对项设置扩展属性，则此类更改可能需要几个小时才能同步。我们建议您改为使用[CustomProperties](/javascript/api/outlook/office.customproperties)对象和相关 api 以避免此类延迟。</span><span class="sxs-lookup"><span data-stu-id="256f0-140">However, if the add-in uses REST or EWS operations to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="256f0-141">若要了解详细信息，请参阅 "在 Outlook 外接程序中获取和设置元数据" 一文中的 "[自定义属性" 部分](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties)。</span><span class="sxs-lookup"><span data-stu-id="256f0-141">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="256f0-142">配置清单</span><span class="sxs-lookup"><span data-stu-id="256f0-142">Configure the manifest</span></span>

<span data-ttu-id="256f0-143">若要在外接程序中启用代理访问方案，必须在父元素`DesktopFormFactor`下的`true`清单中将[SupportsSharedFolders](../reference/manifest/supportssharedfolders.md)元素设置为。</span><span class="sxs-lookup"><span data-stu-id="256f0-143">To enable delegate access scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="256f0-144">目前，其他外观因素不受支持。</span><span class="sxs-lookup"><span data-stu-id="256f0-144">At present, other form factors are not supported.</span></span>

<span data-ttu-id="256f0-145">下面的示例演示`SupportsSharedFolders` `true`在清单的部分中设置的元素。</span><span class="sxs-lookup"><span data-stu-id="256f0-145">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate"></a><span data-ttu-id="256f0-146">将操作作为代理执行</span><span class="sxs-lookup"><span data-stu-id="256f0-146">Perform an operation as delegate</span></span>

<span data-ttu-id="256f0-147">可以通过调用[getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)方法，在撰写或阅读模式下获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="256f0-147">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="256f0-148">这将返回一个[SharedProperties](/javascript/api/outlook/office.sharedproperties)对象，该对象当前提供代理的权限、所有者的电子邮件地址、REST API 的基 URL 和目标邮箱。</span><span class="sxs-lookup"><span data-stu-id="256f0-148">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the delegate's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="256f0-149">下面的示例展示了如何获取邮件或约会的共享属性、检查代理是否具有**写入**权限，以及如何发出 REST 调用。</span><span class="sxs-lookup"><span data-stu-id="256f0-149">The following example shows how to get the shared properties of a message or appointment, check if the delegate has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

## <a name="see-also"></a><span data-ttu-id="256f0-150">另请参阅</span><span class="sxs-lookup"><span data-stu-id="256f0-150">See also</span></span>

- [<span data-ttu-id="256f0-151">允许其他人管理您的邮件和日历</span><span class="sxs-lookup"><span data-stu-id="256f0-151">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="256f0-152">Office 365 中的日历共享</span><span class="sxs-lookup"><span data-stu-id="256f0-152">Calendar sharing in Office 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="256f0-153">如何对清单元素进行排序</span><span class="sxs-lookup"><span data-stu-id="256f0-153">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="256f0-154">[掩码（计算）](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="256f0-154">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="256f0-155">JavaScript 按位运算符</span><span class="sxs-lookup"><span data-stu-id="256f0-155">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)