---
title: 在 Outlook 外接程序中实现追加发送
description: 了解如何在 Outlook 加载项中实现追加发送功能。
ms.topic: article
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18d3e8300a53d08cf484f14cd4fd05adf6382fe3
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541141"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>在 Outlook 外接程序中实现追加发送

本演练结束时，你将拥有一个 Outlook 加载项，该加载项可以在发送消息时插入免责声明。

> [!NOTE]
> 要求集 1.9 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 门，使用 Office 外接程序的 Yeoman 生成器创建加载项项目。

> [!NOTE]
> 如果要使用 [Office 加载项的 Teams 清单 (预览)](../develop/json-manifest-overview.md)，请在 Outlook 快速入门中完成备用快速入门，其中 [包含 Teams 清单 (预览)](../quickstarts/outlook-quickstart-json-manifest.md)，但请在 **“试用”** 部分后跳过所有部分。

## <a name="configure-the-manifest"></a>配置清单

若要配置清单，请打开所使用的清单类型的选项卡。

# <a name="xml-manifest"></a>[XML 清单](#tab/xmlmanifest)

若要在外接程序中启用追加发送功能，必须在 [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) 的集合中包含`AppendOnSend`该权限。

对于此方案，你将运行该函数，而不是在选择 **“执行操作**”按钮时运行`action`函`appendOnSend`数。

1. 在代码编辑器中，打开快速启动项目。

1. 打开位于项目根 **目录的manifest.xml** 文件。

1. 选择整个 **\<VersionOverrides\>** 节点 (包括打开和关闭标记) 并将其替换为以下 XML。

    ```XML
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.9">
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

# <a name="teams-manifest-developer-preview"></a>[Teams 清单 (开发人员预览) ](#tab/jsonmanifest)

1. 打开 manifest.json 文件。

1. 将以下对象添加到“extensions.runtimes”数组。 对于此代码，请注意以下事项。

   - 邮箱要求集的“minVersion”设置为“1.9”，因此无法在不支持此功能的平台和 Office 版本上安装外接程序。 
   - 运行时的“id”设置为描述性名称“function_command_runtime”。
   - “code.page”属性设置为将加载函数命令的无 UI HTML 文件的 URL。
   - “lifetime”属性设置为“short”，这意味着运行时在选择函数命令按钮时启动，并在函数完成时关闭。  (在某些情况下，运行时会在处理程序完成之前关闭。 请参阅 [Office 加载项.) 中的运行时](../testing/runtimes.md)
   - 有一个操作用于运行名为“appendDisclaimerOnSend”的函数。 将在后面的步骤中创建此函数。

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.9"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "function_command_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "appendDisclaimerOnSend",
                "type": "executeFunction",
                "displayName": "appendDisclaimerOnSend"
            }
        ]
    }
    ```

1. 在“authorization.permissions.resourceSpecific”数组中，添加以下对象。 请确保它与数组中具有逗号的其他对象分开。

    ```json
    {
      "name": "Mailbox.AppendOnSend.User",
      "type": "Delegated"
    }
    ```

---

> [!TIP]
> 若要详细了解 Outlook 外接程序的清单，请参阅 [Outlook 加载项清单](manifests.md)。

## <a name="implement-append-on-send-handling"></a>实现追加发送处理

接下来，在发送事件上实现追加。

> [!IMPORTANT]
> 如果外接程序还[使用`ItemSend`它实现发送事件处理](outlook-on-send-addins.md)，则在发送处理程序中调用`AppendOnSendAsync`会返回错误，因为不支持此方案。

对于此方案，你将在用户发送时实现向项目追加免责声明。

1. 在同一快速入门项目中，在代码编辑器中打开文件 **./src/commands/commands.js** 。

1. 在函数之后 `action` ，插入以下 JavaScript 函数。

    ```js
    function appendDisclaimerOnSend(event) {
      const appendText =
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

    ![Outlook 网页版发送时追加免责声明的示例消息。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>另请参阅

[Outlook 加载项清单](manifests.md)
