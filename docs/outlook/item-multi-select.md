---
title: '在多条消息上激活 Outlook 加载项 (预览版) '
description: 了解如何在选择多个消息时激活 Outlook 加载项。
ms.topic: article
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2b77772aa2fc661e4be84c48555e3ddceda224c4
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546436"
---
# <a name="activate-your-outlook-add-in-on-multiple-messages-preview"></a>在多条消息上激活 Outlook 加载项 (预览版) 

使用项目多选功能，Outlook 外接程序现在可以一次性激活并执行对多个选定消息的操作。 某些操作（例如将消息上传到客户关系管理 (CRM) 系统或对许多项进行分类）现在只需单击一下即可轻松完成。

以下部分介绍如何配置外接程序以在读取模式下检索多个消息的主题行。

> [!IMPORTANT]
> 项目多选功能仅在预览版中使用 Outlook on Windows 中的 Microsoft 365 订阅。 预览版中的功能不应用于生产外接程序。我们邀请你在测试或开发环境中测试此功能，并通过 GitHub 欢迎你获得有关体验的反馈 (请参阅本页末尾的 **“反馈** ”部分) 。

> [!NOTE]
> [Teams 清单 (预览版) ](../develop/json-manifest-overview.md)当前不支持项多选功能，但功能组正在努力使此功能可用。

## <a name="prerequisites-to-preview-item-multi-select"></a>预览项目多选的先决条件

若要预览多选功能，请从版本 2209 (Build 15629.20110) 开始安装 Outlook on Windows。 安装后，加入 [Office 预览体验计划](https://insider.office.com/join/windows) ，然后选择 **Beta 通道** 选项以访问 Office beta 生成。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 门，使用 Office 外接程序的 [Yeoman 生成器](../develop/yeoman-generator-overview.md)创建加载项项目。

## <a name="configure-the-manifest"></a>配置清单

若要使外接程序能够在多个选定消息上激活，必须将 [SupportsMultiSelect](/javascript/api/manifest/action?view=outlook-js-preview&preserve-view=true#supportsmultiselect-preview) 子元素添加到 **\<Action\>** 元素并将其值设置为 `true`。 由于项目多选目前仅支持消息， **\<ExtensionPoint\>** 因此元素的 `xsi:type` 属性值必须设置为 `MessageReadCommandSurface` 或 `MessageComposeCommandSurface`设置为 。

1. 在首选代码编辑器中，打开创建的 Outlook 快速入门项目。

1. 打开位于项目根目录的 **manifest.xml** 文件。

1. 为元素`ReadWriteMailbox`分配 **\<Permissions\>** 值。

    ```xml
    <Permissions>ReadWriteMailbox</Permissions>
    ```

1. 选择整个 **\<VersionOverrides\>** 节点，并将其替换为以下 XML。

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
        <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
            <Requirements>
                <bt:Sets DefaultMinVersion="1.12">
                  <bt:Set Name="Mailbox"/>
                </bt:Sets>
            </Requirements>
            <Hosts>
                <Host xsi:type="MailHost">
                    <DesktopFormFactor>
                        <!-- Message Read mode-->
                        <ExtensionPoint xsi:type="MessageReadCommandSurface">
                            <OfficeTab id="TabDefault">
                                <Group id="msgReadGroup">
                                    <Label resid="GroupLabel"/>
                                    <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                        <Label resid="TaskpaneButton.Label"/>
                                        <Supertip>
                                            <Title resid="TaskpaneButton.Label"/>
                                            <Description resid="TaskpaneButton.Tooltip"/>
                                        </Supertip>
                                        <Icon>
                                            <bt:Image size="16" resid="Icon.16x16"/>
                                            <bt:Image size="32" resid="Icon.32x32"/>
                                            <bt:Image size="80" resid="Icon.80x80"/>
                                        </Icon>
                                        <Action xsi:type="ShowTaskpane">
                                            <SourceLocation resid="Taskpane.Url"/>
                                            <!-- Enables your add-in to activate on multiple selected messages. -->
                                            <SupportsMultiSelect>true</SupportsMultiSelect>
                                        </Action>
                                    </Control>
                                </Group>
                            </OfficeTab>
                        </ExtensionPoint>
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
                  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                  <bt:String id="GroupLabel" DefaultValue="Item Multi-select"/>
                  <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                  <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane which displays an option to retrieve the subject line of selected messages."/>
                </bt:LongStrings>
            </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

1. 保存所做的更改。

## <a name="configure-the-task-pane"></a>配置任务窗格

项目多选依赖于 [SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) 事件来确定何时选择或取消选择消息。 此事件需要任务窗格实现。

1. 从 **./src/taskpane** 文件夹中打开 **taskpane.html**。

1. 在元素中 **\<script\>** ，将属性设置 `src` 为 `"https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"`。 这会引用内容分发网络上的 beta 库 (CDN) 。

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"></script>
    ```

1. 在元素中 **\<body\>** ，将整个 **\<main\>** 元素替换为以下标记。

    ```html
    <main id="app-body" class="ms-welcome__main">
        <h2 class="ms-font-xl">Retrieve the subject line of multiple messages with one click!</h2>
        <ul id="selected-items"></ul>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. 保存所做的更改。

## <a name="implement-a-handler-for-the-selecteditemschanged-event"></a>实现 SelectedItemsChanged 事件的处理程序

若要在事件发生时 `SelectedItemsChanged` 发出加载项警报，必须使用 `addHandlerAsync` 该方法注册事件处理程序。

1. 从 **./src/taskpane** 文件夹中打开 **taskpane.js**。

1. 在 `Office.onReady()` 回调函数中，将现有代码替换为以下内容：

    ```javascript
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    
        // Register an event handler to identify when messages are selected.
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }
    
          console.log("Event handler added.");
        });
    }
    ```

## <a name="retrieve-the-subject-line-of-selected-messages"></a>检索所选消息的主题行

注册事件处理程序后，调用 [getSelectedItemsAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getselecteditemsasync-member(1)) 方法以检索所选消息的主题行并将其记录到任务窗格。 该 `getSelectedItemsAsync` 方法还可用于获取其他消息属性，例如项目 ID、项目类型 (`Message` 是目前唯一支持的类型) ，以及项目模式 (`Read` 或 `Compose`) 。

1. 在 **taskpane.js** 中，导航到函 `run` 数并插入以下代码。

    ```javascript
    // Clear list of previously selected messages, if any.
    const list = document.getElementById("selected-items");
    while (list.firstChild) {
        list.removeChild(list.firstChild);
    }

    // Retrieve the subject line of the selected messages and log it to a list in the task pane.
    Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;      
        }

        asyncResult.value.forEach(item => {
            const listItem = document.createElement("li");
            listItem.textContent = item.subject;
            list.appendChild(listItem);
        });
    });
    ```

1. 保存所做的更改。

## <a name="try-it-out"></a>试用

1. 在终端中，在项目的根目录中运行以下代码。 这将启动本地 Web 服务器并旁加载加载项。

    ```command line
    npm start
    ```

    > [!TIP]
    > 如果加载项未自动旁加载，请按照 [Sideload Outlook 加载项中的说明进行测试](sideload-outlook-add-ins-for-testing.md?tabs=windows#outlook-on-the-desktop) ，以便在 Outlook 中手动旁加载它。

1. 在 Outlook 中，确保已启用阅读窗格。 若要启用阅读窗格，请参阅 [“使用”并配置“阅读窗格”以预览消息](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0)。

1. 导航到收件箱，在选择邮件时按 **住 Ctrl** 选择多个消息。

1. 从功能区中选择 **“显示任务窗格** ”。

1. 在任务窗格中，选择 **“运行** ”以查看所选邮件的主题行列表。

    :::image type="content" source="../images/outlook-multi-select.png" alt-text="从多个选定消息检索的主题行的示例列表。":::

## <a name="item-multi-select-behavior-and-limitations"></a>项目多选行为和限制

项目多选仅支持在读取和撰写模式下的 Exchange 邮箱中的消息。 仅当满足以下条件时，Outlook 外接程序才会在多个消息上激活。

- 必须一次从一个 Exchange 邮箱中选择邮件。 不支持非 Exchange 邮箱。
- 必须一次从一个邮箱文件夹中选择邮件。 如果多个邮件位于不同的文件夹中，则加载项不会激活，除非启用了对话视图。 有关详细信息，请参阅 [对话中的多选](#multi-select-in-conversations)。
- 加载项必须实现任务窗格才能检测 `SelectedItemsChanged` 事件。
- 必须启用 Outlook 中的 [阅读窗格](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) 。
- 一次最多可选择 100 条消息。

> [!NOTE]
> 会议邀请和响应被视为消息，而不是约会，因此可以包含在选定内容中。

### <a name="multi-select-in-conversations"></a>对话中的多选

项目多选支持 [对话视图](https://support.microsoft.com/office/0eeec76c-f59b-4834-98e6-05cfdfa9fb07) ，无论它是在邮箱上还是在特定文件夹上启用。 下表描述了展开或折叠对话、选择对话标头时以及对话消息位于与当前查看的文件夹不同的文件夹中时的预期行为。

|选择|展开的对话视图|折叠对话视图|
|------|------|------|
|**已选择对话标头**|如果对话标头是唯一选择的项目，则支持多选的加载项不会激活。 但是，如果还选择了其他非标头消息，则加载项将仅激活这些消息，而不会激活所选标头。|最新消息 (即，会话堆栈中的第一条消息) 包含在消息选择中。<br><br>如果对话中的最新消息位于当前视图中的另一个文件夹中，则位于当前文件夹中的堆栈中的后续消息将包含在所选内容中。|
|**所选对话消息与当前查看的文件夹位于同一文件夹中**|所选的所有对话消息都包含在所选内容中。|不适用。 只有会话标头可用于在折叠对话视图中进行选择。|
|**所选会话消息位于当前视图中的不同文件夹中** |所选的所有对话消息都包含在所选内容中。|不适用。 只有会话标头可用于在折叠对话视图中进行选择。|

## <a name="next-steps"></a>后续步骤

现在，你已启用外接程序以对多个选定消息进行操作，可以扩展外接程序的功能并进一步增强用户体验。 通过将所选邮件的项目 ID 与 [Exchange Web Services (EWS) ](web-services.md) 和 [Microsoft Graph](/graph/overview) 等服务结合使用，探索如何执行更复杂的操作。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项清单](manifests.md)
- [从 Outlook 外接程序调用 Web 服务](web-services.md)
- [Microsoft Graph 概述](/graph/overview)
