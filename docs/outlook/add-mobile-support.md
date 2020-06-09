---
title: 添加对 Outlook 外接程序的移动支持
description: 添加对 Outlook Mobile 的支持需要更新外接程序清单，并且可能会更改移动方案的代码。
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: a4fb02fee8bb429d0193903ba03fcee17b7ede48
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607615"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a>添加对适用于 Outlook Mobile 的外接程序命令的支持

使用 Outlook Mobile 中的外接程序命令，用户可以访问在 web、Windows 和 Mac 上的 Outlook 中已有的相同功能（有一些[限制](#code-considerations)）。 添加对 Outlook Mobile 的支持需要更新外接程序清单，并且可能会更改移动方案的代码。

## <a name="updating-the-manifest"></a>更新清单

启用 Outlook Mobile 中的外接程序命令的第一步是在外接程序清单中对其进行定义。[VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 架构为移动电话定义新的外形规格，[MobileFormFactor](../reference/manifest/mobileformfactor.md)。

此元素包含在移动客户端中加载外接程序所需的所有信息。这使你可以为移动体验定义完全不同的 UI 元素和 JavaScript 文件。

下面的示例显示元素中的单个任务窗格按钮 `MobileFormFactor` 。

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

这与 [DesktopFormFactor](../reference/manifest/desktopformfactor.md) 元素中出现的元素非常相似，但具有一些明显的区别。

- 不使用 [OfficeTab](../reference/manifest/officetab.md) 元素。
- [ExtensionPoint](../reference/manifest/extensionpoint.md) 元素必须仅具有一个子元素。如果外接程序仅添加一个按钮，则子元素应为 [Control](../reference/manifest/control.md) 元素。如果外接程序添加多个按钮，则子元素应为包含多个 `Control` 元素的 [Group](../reference/manifest/group.md) 元素。
- 没有与 `Control` 元素等效的 `Menu` 类型。
- 不使用 [Supertip](../reference/manifest/supertip.md) 元素。
- 要求的图标大小不同。移动外接程序最少必须支持 25x25、32x32 和 48x48 像素的图标。

## <a name="code-considerations"></a>代码注意事项

设计适用于移动电话的外接程序引入了一些额外注意事项。

### <a name="use-rest-instead-of-exchange-web-services"></a>使用 REST 代替 Exchange Web 服务

Outlook Mobile 中不支持 [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法。外接程序应在可能的情况下首选从 Office.js API 获取信息。如果外接程序需要 Office.js API 未公开的信息，则应使用 [Outlook REST API](/outlook/rest/) 来访问用户邮箱。

邮箱要求集1.5 引入了新版本的[mailbox.getcallbacktokenasync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) ，该版本可以请求与 REST api 兼容的访问令牌，以及可用于查找用户的 rest api 终结点的新的[office.context.mailbox.resturl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)属性。

### <a name="pinch-zoom"></a>收缩缩放

在默认情况下，用户可以使用“收缩缩放”手势在任务窗格上进行缩放。如果方案不需要该功能，请确保在 HTML 中禁用收缩缩放。

### <a name="close-task-panes"></a>关闭任务窗格

在 Outlook Mobile 中，任务窗格占据整个屏幕，并且在默认情况下需要用户将其关闭以返回到邮件。请考虑使用 [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) 方法在方案完成时关闭任务窗格。

### <a name="compose-mode-and-appointments"></a>撰写模式和约会

目前，Outlook Mobile 中的外接程序仅在读取邮件时支持激活。 在撰写邮件时或查看或撰写约会时，不会激活外接程序。 但是，可以在约会组织者模式下激活联机会议提供程序集成的外接程序。 有关此异常的详细信息，请参阅[创建适用于联机会议提供商文章的 Outlook mobile 外](online-meeting.md)接程序。

### <a name="unsupported-apis"></a>不支持的 API

Outlook Mobile 不支持在要求集1.6 或更高版本中引入的 Api。 此外，还不支持来自早期要求集的以下 Api。

  - [Office.context.officeTheme](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
  - [Office.context.mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
  - [Office.context.mailbox.convertToEwsId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
  - [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
  - [Office.context.mailbox.item.displayReplyAllForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.displayReplyForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
  - [Office.context.mailbox.item.getRegexMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a>另请参阅

[要求集支持](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)