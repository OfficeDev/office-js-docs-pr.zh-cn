---
title: 在Outlook加载项中实现追加发送
description: 了解如何在Outlook加载项中实现追加发送功能。
ms.topic: article
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 968b730aca1fc36640e43ff45404c8d4c7b92d47
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244833"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>在Outlook加载项中实现追加发送

本演练结束时，你将有一个Outlook加载项，该加载项可以在发送消息时插入免责声明。

> [!NOTE]
> 要求集 1.9 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="set-up-your-environment"></a>设置环境

完成[Outlook快速入](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)门，使用 yeoman 生成器为Office加载项创建加载项项目。

## <a name="configure-the-manifest"></a>配置清单

若要在外接程序中启用追加发送功能，必须在 [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) 的集合中包含`AppendOnSend`该权限。

对于此方案，你将运行该函数，而不是在选择 **“执行操作**”按钮时运行`action`函`appendOnSend`数。

1. 在代码编辑器中，打开快速启动项目。

1. 打开位于项目根 **目录的manifest.xml** 文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括打开和关闭标记) 并将其替换为以下 XML。

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
> 若要详细了解Outlook加载项的清单，请[参阅Outlook加载项清单](manifests.md)。

## <a name="implement-append-on-send-handling"></a>实现追加发送处理

接下来，在发送事件上实现追加。

> [!IMPORTANT]
> 如果外接程序还[使用`ItemSend`它实现发送事件处理](outlook-on-send-addins.md)，则在发送处理程序中调用`AppendOnSendAsync`会返回错误，因为不支持此方案。

对于此方案，你将在用户发送时实现向项目追加免责声明。

1. 在同一快速入门项目中，在代码编辑器中打开文件 **./src/commands/commands.js** 。

1. 在函数之后 `action` ，插入以下 JavaScript 函数。

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
    
1. 函数下方立即添加以下行以注册函数。

    ```js
    Office.actions.associate("appendDisclaimerOnSend", appendDisclaimerOnSend);
    ```

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 运行此命令时，如果本地 Web 服务器尚未运行，并且外接程序将旁加载，则会启动该服务器。 

    ```command&nbsp;line
    npm start
    ```

1. 创建新消息，并将其添加到 **To** 行。

1. 在功能区或溢出菜单中，选择 **“执行操作**”。

1. 发送邮件，然后从 **收件箱** 或 **“已发送邮件”** 文件夹中打开邮件以查看追加的免责声明。

    ![示例消息的屏幕截图，其中附有在发送Outlook 网页版中的免责声明。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>另请参阅

[Outlook 加载项清单](manifests.md)
