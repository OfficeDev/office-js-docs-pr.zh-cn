---
title: 验证并排查清单问题
description: 使用这些方法验证 Office 加载项清单。
ms.date: 11/02/2018
localization_priority: Priority
ms.openlocfilehash: 921adf6f1f398887d96031790facc1fb1425af2b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451150"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>验证并排查清单问题

下面这些方法可用于验证和排查 Office 加载项清单问题： 

- [通过 Office 加载项验证程序验证清单](#validate-your-manifest-with-the-office-add-in-validator)   
- [根据 XML 架构验证清单](#validate-your-manifest-against-the-xml-schema)
- [使用适用于 Office 加载项的 Yeoman 生成器来验证清单](#validate-your-manifest-with-the-yeoman-generator-for-office-add-ins)
- [使用运行时日志记录功能调试加载项](#use-runtime-logging-to-debug-your-add-in)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>通过 Office 加载项验证程序验证清单

为了确保描述 Office 加载项的清单文件正确完整，请使用 [Office 加载项验证程序](https://github.com/OfficeDev/office-addin-validator)验证清单。

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a>使用 Office 加载项验证程序验证清单的具体步骤

1. 安装 [Node.js](https://nodejs.org/download/)。 

2. 以管理员身份打开命令提示符/终端，并运行下面的命令，以全局安装 Office 加载项验证程序及其依赖项：

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > 如果已安装 Yo Office，请升级到最新版本，验证程序就会作为依赖项进行安装。

3. 运行下面的命令来验证清单。将 MANIFEST.XML 替换为清单 XML 文件路径。

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>根据 XML 架构验证清单

为了有助于确保清单文件采用正确架构，请为要使用的元素添加任何命名空间。 如果从其他示例清单中复制了元素，请仔细检查是否还**添加了相应命名空间**。 可以根据 [XML 架构定义 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件验证清单。 若要执行此验证，可以使用 XML 架构验证工具。 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>使用命令行 XML 架构验证工具验证清单的具体步骤

1.  安装 [tar](https://www.gnu.org/software/tar/) 和 [libxml](http://xmlsoft.org/FAQ.html)（如果尚未安装的话）。

2.  运行下面的命令。将 `XSD_FILE` 替换为清单 XSD 文件路径，并将 `XML_FILE` 替换为清单 XML 文件路径。
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>使用适用于 Office 加载项的 Yeoman 生成器来验证清单

如果已使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)创建 Office 加载项，则可以通过在项目的根目录中运行以下命令来确保清单文件遵循正确的架构：

```bash
npm run validate
```

![动画 gif 显示 Yo Office 验证程序（在命令行处运行并生成显示“验证已通过”的结果）。](../images/yo-office-validator.gif)

> [!NOTE]
> 若要访问此功能，必须使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)版本 1.1.17 或更高版本创建加载项项目。

## <a name="use-runtime-logging-to-debug-your-add-in"></a>使用运行时日志记录功能调试加载项 

可以使用运行时日志记录调试加载项的清单以及多个安装错误。 此功能可以帮助你标识并修复清单中未被 XSD 架构验证检测到的问题，例如资源 ID 间的不匹配等。 运行时日志记录对于调试执行加载项命令的加载项和 Excel 自定义功能尤其有用。   

> [!NOTE]
> 运行时日志记录功能暂适用于 Office 2016 桌面版。

### <a name="to-turn-on-runtime-logging"></a>启用运行时日志记录功能的具体步骤

> [!IMPORTANT]
> 运行时日志记录功能影响性能。仅在需要调试加载项清单问题时，才启用此功能。

若要启用运行时日志记录功能，请执行以下操作：

1. 确保运行的是 Office 2016 桌面版 **16.0.7019** 或更高版本。 

2. 在 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` 下添加 `RuntimeLogging` 注册表项。 

    > [!NOTE]
    > 如果 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\` 下尚不存在 `Developer` 密钥（文件夹），请完成以下步骤以创建它： 
    > 1. 右键单击 **WEF** 密钥（文件夹），然后选择**新建** > **密钥**。
    > 2. 将新密钥命名为 **Developer**。

3. 将此项的默认值设置为你想要在其中写入日志的文件的完整路径。有关示例，请参阅 [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip)。 

    > [!NOTE]
    > 向其中写入日志文件的目录必须已经存在，并且必须拥有对它的写入权限。 
 
注册表应如下图所示。 若要禁用此功能，请从注册表中删除 `RuntimeLogging`。 

![包含 RuntimeLogging 注册表项的注册表编辑器屏幕截图](http://i.imgur.com/Sa9TyI6.png)


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
删除文件夹 `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` 的内容。

#### <a name="for-ios"></a>对于 iOS：
在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。

## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](../develop/add-in-manifests.md)
- [旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [调试 Office 外接程序](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
