---
title: 在 Outlook 外接程序中实施 "在发送时追加" （预览）
description: 了解如何在 Outlook 外接程序中实现 "发送时发送" 功能。
ms.topic: article
ms.date: 05/26/2020
localization_priority: Normal
ms.openlocfilehash: f7f345ad726529c7ba3f8fa3ceedb46246310547
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607594"
---
# <a name="implement-append-on-send-in-your-outlook-add-in-preview"></a>在 Outlook 外接程序中实施 "在发送时追加" （预览）

本演练结束时，您将拥有一个可在发送邮件时插入免责声明的 Outlook 外接程序。

> [!IMPORTANT]
> 此功能目前仅支持在 Outlook 网页版和使用 Office 365 订阅的 Windows 中进行[预览](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。 有关更多详细信息，请参阅[如何预览本文中的追加发送功能](#how-to-preview-the-append-on-send-feature)。
>
> 由于预览功能可能会发生更改，恕不另行通知，它们不应在生产外接程序中使用。

## <a name="how-to-preview-the-append-on-send-feature"></a>如何预览追加发送功能

我们邀请你试用 "发送时追加" 功能！ 让我们知道你的方案以及我们如何通过 GitHub 向我们提供反馈（请参阅本页结尾处的**反馈**部分）来改进你的情况。

若要预览此功能：

- 参考 CDN 上的**beta**库（ https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 。 在 CDN 和[jquery.typescript.definitelytyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)中找到 TypeScript 编译和智能感知的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。 您可以使用安装这些类型 `npm install --save-dev @types/office-js-preview` 。
- 对于 Windows，你可能需要加入[Office 预览体验成员计划](https://insider.office.com)，以访问更多最近的 office 版本。
- 对于 web 上的 Outlook，[在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)。

## <a name="set-up-your-environment"></a>设置环境

完成[Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)，它将使用 Office 外接程序的 Yeoman 生成器创建外接程序项目。

## <a name="configure-the-manifest"></a>配置清单

若要在您的外接程序中启用 "追加发送" 功能，必须 `AppendOnSend` 在[ExtendedPermissions](../reference/manifest/extendedpermissions.md)集合中包含该权限。

对于此方案， `action` 您将运行函数，而不是在选择 "**执行操作**" 按钮时运行函数 `appendOnSend` 。

1. 在代码编辑器中，打开 "快速启动" 项目。

1. 打开位于项目根目录的**清单 .xml**文件。

1. 选择整个 `<VersionOverrides>` 节点（包括 "打开" 和 "关闭" 标记），并将其替换为以下 XML。

    ```XML
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.3">
            <bt:Set Name="Mailbox" />
          </bt:Sets>
        </Requirements>
        <Hosts>
          <Host xsi:type="MailHost">
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
                      <Label resid="TaskpaneButton.Label" />
                      <Supertip>
                        <Title resid="TaskpaneButton.Label" />
                        <Description resid="TaskpaneButton.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16x16" />
                        <bt:Image size="32" resid="Icon.32x32" />
                        <bt:Image size="80" resid="Icon.80x80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <SourceLocation resid="Taskpane.Url" />
                      </Action>
                    </Control>
                    <Control xsi:type="Button" id="ActionButton">
                      <Label resid="ActionButton.Label"/>
                      <Supertip>
                        <Title resid="ActionButton.Label"/>
                        <Description resid="ActionButton.Tooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16x16"/>
                        <bt:Image size="32" resid="Icon.32x32"/>
                        <bt:Image size="80" resid="Icon.80x80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

            </DesktopFormFactor>
          </Host>
        </Hosts>
        <Resources>
          <bt:Images>
            <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
            <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
            <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
            <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
            <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
            <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
            <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
            <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
            <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
          </bt:LongStrings>
        </Resources>
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> 若要了解有关 Outlook 外接程序的清单的详细信息，请参阅[outlook 外接程序清单](manifests.md)。

## <a name="implement-append-on-send-handling"></a>实现附加发送前处理

接下来，实现在 send 事件上追加。

在这种情况下，您将实现在用户发送时向项目追加免责声明。

1. 在同一 "快速启动" 项目中，在代码编辑器中打开 **/src/commands/commands.js** 。

1. 在 `action` 函数后面，插入以下 JavaScript 函数。

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. 在文件末尾，添加以下语句。

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 运行此命令时，本地 web 服务器将启动（如果它尚未运行）。

    ```command&nbsp;line
    npm run dev-server
    ```

1. 按照[旁加载 Outlook 外接程序](sideload-outlook-add-ins-for-testing.md)中的说明进行操作，以进行测试。

1. 创建新邮件，并将自己添加到 " **to** " 行。

1. 从 "功能区" 或 "溢出" 菜单中，选择 "**执行操作**"。

1. 发送邮件，然后从 **"收件箱" 或 "** **已发送邮件**" 文件夹中打开它以查看追加的免责声明。

    ![在 Outlook 网页版上追加的包含免责声明的示例邮件的屏幕截图。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>另请参阅

[Outlook 加载项清单](manifests.md)