---
title: 添加对 Outlook 外接程序的移动支持
description: 添加对 Outlook Mobile 的支持需要更新外接程序清单，并且可能会更改移动方案的代码。
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6e99c862d4cd63590a86c757bf2b720c096826a9
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496969"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>添加对适用于 Outlook Mobile 的外接程序命令的支持

使用 Outlook Mobile 中的外接程序命令，用户可以访问与 (、) 、Outlook 网页版、Windows 和 Mac 中已有的相同功能[](#code-considerations)。 添加对 Outlook Mobile 的支持需要更新外接程序清单，并且可能会更改移动方案的代码。

## <a name="updating-the-manifest"></a>更新清单

启用 Outlook Mobile 中的外接程序命令的第一步是在外接程序清单中对其进行定义。[VersionOverrides](/javascript/api/manifest/versionoverrides) v1.1 架构为移动电话定义新的外形规格，[MobileFormFactor](/javascript/api/manifest/mobileformfactor)。

此元素包含在移动客户端中加载外接程序所需的所有信息。这使你可以为移动体验定义完全不同的 UI 元素和 JavaScript 文件。

以下示例显示了元素中的单个任务窗格 `MobileFormFactor` 按钮。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

这与 [DesktopFormFactor](/javascript/api/manifest/desktopformfactor) 元素中出现的元素非常相似，但具有一些明显的区别。

- 不使用 [OfficeTab](/javascript/api/manifest/officetab) 元素。
- [ExtensionPoint](/javascript/api/manifest/extensionpoint) 元素必须仅具有一个子元素。如果外接程序仅添加一个按钮，则子元素应为 [Control](/javascript/api/manifest/control) 元素。如果外接程序添加多个按钮，则子元素应为包含多个 `Control` 元素的 [Group](/javascript/api/manifest/group) 元素。
- 没有与 `Control` 元素等效的 `Menu` 类型。
- 不使用 [Supertip](/javascript/api/manifest/supertip) 元素。
- 要求的图标大小不同。移动外接程序最少必须支持 25x25、32x32 和 48x48 像素的图标。

## <a name="code-considerations"></a>代码注意事项

设计适用于移动电话的外接程序引入了一些额外注意事项。

### <a name="use-rest-instead-of-exchange-web-services"></a>使用 REST 代替 Exchange Web 服务

Outlook Mobile 中不支持 [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法。外接程序应在可能的情况下首选从 Office.js API 获取信息。如果外接程序需要 Office.js API 未公开的信息，则应使用 [Outlook REST API](/outlook/rest/) 来访问用户邮箱。

邮箱要求集 1.5 引入了可请求与 REST API 兼容的访问令牌的 [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 新版本，以及可用于查找用户的 REST API 终结点的新 [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) 属性。

### <a name="pinch-zoom"></a>收缩缩放

在默认情况下，用户可以使用“收缩缩放”手势在任务窗格上进行缩放。如果方案不需要该功能，请确保在 HTML 中禁用收缩缩放。

### <a name="close-task-panes"></a>关闭任务窗格

在 Outlook Mobile 中，任务窗格占据整个屏幕，并且在默认情况下需要用户将其关闭以返回到邮件。请考虑使用 [Office.context.ui.closeContainer](/javascript/api/office/office.ui#office-office-ui-closecontainer-member(1)) 方法在方案完成时关闭任务窗格。

### <a name="compose-mode-and-appointments"></a>撰写模式和约会

目前，移动版中的Outlook仅在阅读邮件时支持激活。 在撰写邮件时或查看或撰写约会时，不会激活外接程序。 但是，联机会议提供商集成加载项可以在约会管理器模式下激活。 有关此异常 (包括可用的 API) ，请参阅为联机会议Outlook创建移动[外接程序。](online-meeting.md#available-apis)

### <a name="unsupported-apis"></a>不支持的 API

要求集 1.6 或更高版本中引入的 API 不受 Outlook Mobile 支持。 还不支持来自早期要求集的以下 API。

- [Office.context.officeTheme](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context#officetheme-officetheme)
- [Office.context.mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
- [Office.context.mailbox.convertToEwsId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayMessageForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
- [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
- [Office.context.mailbox.item.displayReplyAllForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.displayReplyForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntities](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getEntitiesByType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [Office.context.mailbox.item.getRegexMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

## <a name="see-also"></a>另请参阅

[Exchange 服务器和 Outlook 客户端支持的要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)