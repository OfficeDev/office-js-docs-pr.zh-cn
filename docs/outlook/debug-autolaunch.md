---
title: 调试基于事件的 Outlook 加载项
description: 了解如何调试实现基于事件的激活的 Outlook 加载项。
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5d36a23b34132071077e3eb192e562288befb8a5
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797489"
---
# <a name="debug-your-event-based-outlook-add-in"></a>调试基于事件的 Outlook 加载项

本文提供调试指南，以便在加载项中实现 [基于事件的激活](autolaunch.md) 。 基于事件的激活功能已在 [要求集 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10) 中引入，其他事件现在以预览版提供。 有关详细信息，请参阅 [支持的事件](autolaunch.md#supported-events)。

> [!IMPORTANT]
> 此调试功能仅在具有 Microsoft 365 订阅的 Outlook on Windows 中受支持。

本文介绍启用调试的关键阶段。

- [标记用于调试的加载项](#mark-your-add-in-for-debugging)
- [配置Visual Studio Code](#configure-visual-studio-code)
- [附加Visual Studio Code](#attach-visual-studio-code)
- [Debug](#debug)

例如，如果使用适用于 Office 外接程序的 Yeoman 生成器创建外接程序项目 (，请执行 [基于事件的激活演练](autolaunch.md)) ，然后在本文中遵循 **“使用 Yeoman 创建”生成器** 选项。 否则，请执行 **其他** 步骤。 Visual Studio Code应至少为版本 1.56.1。

## <a name="mark-your-add-in-for-debugging"></a>标记加载项以进行调试

1. 设置注册表项 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`。 `[Add-in ID]`**\<Id\>** 是加载项清单中。

    **使用 Yeoman 生成器创建**：在命令行窗口中，导航到加载项文件夹的根，然后运行以下命令。

    ```command&nbsp;line
    npm start
    ```

    除了生成代码和启动本地服务器外，此命令还应将此加载项的`1`注册表项设置`UseDirectDebugger`为。

    **其他**：在 `UseDirectDebugger` 下面 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`添加注册表项。 替换 `[Add-in ID]` 为 **\<Id\>** 加载项清单中的清单。 将注册表项设置为 `1`.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. 启动 Outlook 桌面 (或重启 Outlook（如果已打开) ）。
1. 撰写新消息或约会。 应会看到以下对话框。 尚 *不* 与对话进行交互。

    ![“调试基于事件的处理程序”对话框的屏幕截图。](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>配置Visual Studio Code

### <a name="created-with-yeoman-generator"></a>使用 Yeoman 生成器创建

1. 返回命令行窗口，打开Visual Studio Code。

    ```command&nbsp;line
    code .
    ```

1. 在Visual Studio Code中，打开 **文件 ./.vscode/launch.json**，并将以下摘录添加到配置列表。 保存所做的更改。

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a>其他

1. 可能在 **桌面** 文件夹) 中创建名为“**调试**”的新文件夹 (。
1. 打开 Visual Studio Code。
1. 转到 **“文件** > **打开文件夹**”，导航到刚创建的文件夹，然后选择 **“选择文件夹**”。
1. 在活动栏上，选择“ **调试** ”项 (Ctrl+Shift+D) 。

    ![活动栏上“调试”图标的屏幕截图。](../images/vs-code-debug.png)

1. 选择 **创建 launch.json 文件** 链接。

    ![用于在Visual Studio Code中创建 launch.json 文件的链接的屏幕截图。](../images/vs-code-create-launch.json.png)

1. 在 **“选择环境** ”下拉列表中，选择 **“边缘：启动** ”以创建 launch.json 文件。
1. 将以下摘录添加到配置列表。 保存所做的更改。

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a>附加Visual Studio Code

1. 若要查找加载项的 **bundle.js**，请在 Windows 资源管理器中打开以下文件夹，并搜索清单) 中找到的 **\<Id\>** 加载项 (。

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    打开以此 ID 为前缀的文件夹并复制其完整路径。 在Visual Studio Code中，打开该文件夹中的 **bundle.js**。 文件路径的模式应如下所示：

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. 将断点置于要在其中停止调试器的bundle.js。
1. 在 **“调试”** 下拉列表中，选择“ **直接调试**”名称，然后选择 **“运行**”。

    ![从“Visual Studio Code调试”下拉列表中的配置选项中选择“直接调试”的屏幕截图。](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>调试

1. 确认调试器已附加后，返回到 Outlook，然后在 **“基于调试事件的处理程序** ”对话框中，选择 **“确定** ”。

1. 现在可以在Visual Studio Code中点击断点，以便调试基于事件的激活代码。

## <a name="stop-debugging"></a>停止调试

若要停止调试当前 Outlook 桌面会话的其余部分，请在 **“基于调试事件的处理程序** ”对话框中选择 **“取消**”。 若要重新启用调试，请重启 Outlook 桌面。

若要防止 **弹出基于调试事件的处理程序** 对话框并停止调试后续 Outlook 会话，请删除关联的注册表项或将其值设置为 `0`： `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`。

## <a name="see-also"></a>另请参阅

- [配置 Outlook 外接程序以进行基于事件的激活](autolaunch.md)
- [使用运行时日志记录功能调试加载项](../testing/runtime-logging.md#runtime-logging-on-windows)
