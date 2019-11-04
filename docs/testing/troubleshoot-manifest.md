---
title: 验证并排查清单问题
description: 使用这些方法验证 Office 加载项清单。
ms.date: 10/29/2019
localization_priority: Priority
ms.openlocfilehash: c1af6308a975bf9204a519e21f828454d286aa19
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924806"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>验证并排查清单问题

你可能需要验证加载项的清单文件，以确保其正确且完整。 当你尝试旁加载加载项时，验证还可以识别导致错误“你的加载项清单无效”的问题。 本文介绍了验证清单文件和解决加载项问题的多种方法。

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>使用适用于 Office 加载项的 Yeoman 生成器来验证清单

如果你使用了[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则也可以使用它来验证项目的清单文件。 在项目的根目录中运行以下命令：

```command&nbsp;line
npm run validate
```

![动画 gif 显示 Yo Office 验证程序（在命令行处运行并生成显示“验证已通过”的结果）。](../images/yo-office-validator.gif)

> [!NOTE]
> 若要访问此功能，必须使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)版本 1.1.17 或更高版本创建加载项项目。

## <a name="validate-your-manifest-with-office-addin-manifest"></a>使用 office-addin-manifest 验证清单

如果你未使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则可以使用 [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest)。

1. 安装 [Node.js](https://nodejs.org/download/)。

2. 在项目的根目录中运行以下命令。 将 `MANIFEST_FILE` 替换为清单文件的名称。

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > 如果运行此命令导致错误消息“命令语法无效。” （因为 `validate` 命令无法识别），运行以下命令验证清单（用清单文件的名称替换 `MANIFEST_FILE`）： 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a>根据 XML 架构验证清单

可以根据 [XML 架构定义 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件来验证清单文件。 这将有助于确保清单文件采用正确架构（包括所使用的元素的所有命名空间）。 如果从其他示例清单中复制了元素，请仔细检查是否还**添加了相应命名空间**。 若要执行此验证，可以使用 XML 架构验证工具。

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>使用命令行 XML 架构验证工具验证清单的具体步骤

1. 安装 [tar](https://www.gnu.org/software/tar/) 和 [libxml](http://xmlsoft.org/FAQ.html)（如果尚未安装的话）。

2. 运行下面的命令。将 `XSD_FILE` 替换为清单 XSD 文件路径，并将 `XML_FILE` 替换为清单 XML 文件路径。
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a>使用运行时日志记录功能调试加载项

可以使用运行时日志记录调试加载项的清单以及多个安装错误。 此功能可以帮助你标识并修复清单中未被 XSD 架构验证检测到的问题，例如资源 ID 间的不匹配等。 运行时日志记录对于调试执行加载项命令的加载项和 Excel 自定义功能尤其有用。   

> [!NOTE]
> 运行时日志记录功能暂适用于 Office 2016 桌面版。

> [!IMPORTANT]
> 运行时日志记录影响性能。 请仅在需要调试外接程序清单中的问题时启用此功能。

### <a name="use-runtime-logging-from-the-command-line"></a>使用命令行中的运行时日志

从命令行启用运行时日志记录是最快的使用此日志记录工具的方式。 这些使用 npx，默认情况下，它作为 npm@5.2.0+ 的一部分提供。 如果使用的是 [npm](https://www.npmjs.com/) 的早期版本，请尝试 [Windows 上的运行时日志记录](#runtime-logging-on-windows)或 [Mac](#runtime-logging-on-mac)说明，或者[安装 npx](https://www.npmjs.com/package/npx)。

- 要启用运行时日志记录，请执行以下操作：
    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```
- 若要仅对特定文件启用运行时日志记录，请使用包含文件名的相同命令：

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- 要禁用运行时日志记录，请执行以下操作：

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- 要显示是否启用了运行时日志记录，请执行以下操作：

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- 要在命令行内显示有关运行时日志记录的帮助，请执行以下操作：

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

### <a name="runtime-logging-on-windows"></a>Windows 上的运行时日志记录

1. 确保运行的是 Office 2016 桌面版 **16.0.7019** 或更高版本。 

2. 在 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` 下添加 `RuntimeLogging` 注册表项。 

    > [!NOTE]
    > 如果 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\` 下尚不存在 `Developer` 密钥（文件夹），请完成以下步骤以创建它： 
    > 1. 右键单击 **WEF** 密钥（文件夹），然后选择**新建** > **密钥**。
    > 2. 将新密钥命名为 **Developer**。

3. 将 **RuntimeLogging** 项的默认值设置为你想要在其中写入日志的文件的完整路径。 有关示例，请参阅 [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip)。 

    > [!NOTE]
    > 向其中写入日志文件的目录必须已经存在，并且必须拥有对它的写入权限。 
 
注册表应如下图所示。 若要禁用此功能，请从注册表中删除 `RuntimeLogging`。 

![包含 RuntimeLogging 注册表项的注册表编辑器屏幕截图](http://i.imgur.com/Sa9TyI6.png)

### <a name="runtime-logging-on-mac"></a>Mac 上的运行时日志记录

1. 确保运行的是 Office 2016 桌面版 **16.27** (19071500) 或更高版本。

2. 打开**终端**并使用 `defaults` 命令设置运行时日志记录首选项：
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    `<bundle id>` 确定了主机要对哪些运行时日志记录。 `<file_name>` 是要将日志写入的文本文件的名称。

    将 `<bundle id>` 设置为下述值之一，从而为相应主机启用运行时日志记录：

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

以下示例会为 Word 启用运行时日志记录，然后打开日志文件：

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> 运行 `defaults` 命令来启用运行时日志记录后，需要重启 Office。

要关闭运行时日志记录，请使用 `defaults delete` 命令：

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

以下示例将为 Word 关闭运行时日志记录：

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

### <a name="to-troubleshoot-issues-with-your-manifest"></a>排查清单问题的具体步骤

若要使用运行时日志记录功能排查加载项的加载问题，请执行以下操作：
 
1. [旁加载加载项](sideload-office-add-ins-for-testing.md)以供测试。 

    > [!NOTE]
    > 建议仅旁加载要测试的加载项，以最大限度地减少日志文件中的消息数。

2. 如果没有任何变化，且看不到加载项（加载项对话框中没有显示），请打开日志文件。

3. 在日志文件中搜索你的外接程序 ID（已在清单中定义）。在日志文件中，此 ID 标有 `SolutionId`。 

在以下示例中，日志文件标识指向某个不存在的资源文件的控件。对于此示例，修复方法则是更正清单中的输入错误或添加丢失的资源。

![带有指定未找到的资源 ID 的条目的日志文件屏幕截图](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>运行时日志记录已知问题

在日志文件中看到的信息可能很混乱或其分类不正确。例如：

- 后跟 `Unexpected Parsed manifest targeting different host` 的信息 `Medium Current host not in add-in's host list` 是错误分类为错误。

- 如果发现信息 `Unexpected Add-in is missing required manifest fields DisplayName` 且其不包含 SolutionId，则此错误极可能与你正在调试的外接程序无关。 

- 对系统而言，任何 `Monitorable` 信息都会视其为错误。有时这些信息表示清单中的问题，例如一个已跳过但未引起清单运行失败的拼写错误的元素。 

## <a name="clear-the-office-cache"></a>清除 Office 缓存

如果在清单中进行的更改（如功能区按钮图标的文件名或加载项命令的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。 

#### <a name="for-windows"></a>对于 Windows：
删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。

#### <a name="for-mac"></a>对于 Mac：

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>对于 iOS：
在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。

## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](../develop/add-in-manifests.md)
- [旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [调试 Office 外接程序](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
