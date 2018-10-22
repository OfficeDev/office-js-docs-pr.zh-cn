---
title: 验证并排查清单问题
description: 使用这些方法验证 Office 加载项清单。
ms.date: 12/04/2017
ms.openlocfilehash: 51d644f7cfb7fbad5c9b66be41dc57015202b9be
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/19/2018
ms.locfileid: "25639985"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>验证并排查清单问题

使用下面这些方法验证和排查 Office 加载项清单问题： 

- [通过 Office 加载项验证程序验证清单](#validate-your-manifest-with-the-office-add-in-validator)   
- [根据 XML 架构验证清单](#validate-your-manifest-against-the-xml-schema)
- [使用运行时日志记录功能调试加载项清单](#use-runtime-logging-to-debug-your-add-in-manifest)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>通过 Office 加载项验证程序验证清单

为了确保描述 Office 加载项的清单文件正确完整，请使用 [Office 加载项验证程序](https://github.com/OfficeDev/office-addin-validator)验证清单。

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a>使用Office加载项验证程序验证清单

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

为了帮助确保清单文件遵循正确的架构，其中包括任何要使用的元素的命名空间。如果复制了其他样本清单中的元素，也请仔细检查 **包括适当的命名空间**。可以根据标准验证清单 [XML 架构定义 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件。可以使用 XML 架构验证工具执行此验证。 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>使用命令行 XML 架构验证工具验证清单

1.  安装 [tar](https://www.gnu.org/software/tar/) 和 [libxml](http://xmlsoft.org/FAQ.html)（如果尚未安装的话）。

2.  运行下面的命令。将 `XSD_FILE` 替换为清单 XSD 文件路径，并将 `XML_FILE` 替换为清单 XML 文件路径。
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a>使用运行时日志记录功能调试加载项 

可以使用运行时日志记录来调试加载项的清单以及一些安装错误。此功能可帮助识别和修复XSD架构验证未检测到的清单问题，例如资源ID不匹配。运行时日志记录对于调试实现加载项命令和 Excel 自定义函数的加载项特别有用。   

> [!NOTE]
> 运行时日志记录功能目前用于 Office 2016 桌面版。

### <a name="to-turn-on-runtime-logging"></a>启用运行时日志记录功能

> [!IMPORTANT]
> 运行时日志记录功能影响性能。仅在需要调试加载项清单问题时，才启用此功能。

启用运行时日志记录功能：

1. 确保运行的是 Office 2016 桌面版 **16.0.7019** 或更高版本。 

2. 添加`RuntimeLogging`注册表项 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\` 

3. 将此项的默认值设置为欲在其中写入日志的文件的完整路径。有关示例，请参阅 [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip)。 

    > [!NOTE]
    > 向其中写入日志文件的目录必须已经存在，并且必须拥有写入权限。 
 
注册表的模样如下图所示。若要关闭该功能，请从注册表中删除 `RuntimeLogging` 项。 

![注册表编辑器的屏幕截图，其中包含 RuntimeLogging 注册表项](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a>排查清单问题

使用运行时日志记录解决加载外接程序的问题：
 
1. [旁加载加载项](sideload-office-add-ins-for-testing.md)以供测试。 

    > [!NOTE]
    > 建议仅旁加载要测试的加载项，以最大限度地减少日志文件中的消息数。

2. 如果没有任何变化，且看不到加载项（加载项对话框中没有显示），请打开日志文件。

3. 在日志文件中搜索在清单中定义的加载项ID。在日志文件中，此 ID 标有 `SolutionId`。 

在以下示例中，日志文件标识指向某个不存在的资源文件的控件。对于此示例，修复方法则是更正清单中的输入错误或添加丢失的资源。

![日志文件的屏幕截图，其中包含指定未找到的资源ID的条目](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>运行时日志记录的已知问题

在日志文件中看到的信息可能很混乱或其分类不正确。例如：

- 信息 `Medium Current host not in add-in's host list` 紧跟 `Unexpected Parsed manifest targeting different host` ，被错误地归类为错误。

- 如果发现信息 `Unexpected Add-in is missing required manifest fields DisplayName` 且其不包含 SolutionId，则此错误极可能与正在调试的加载项无关。 

- 对系统而言，任何 `Monitorable` 信息都会视其为错误。有时它们表明清单存在问题，例如，跳过但未导致清单失败的拼写错误的元素。 

## <a name="clear-the-office-cache"></a>清除 Office 缓存

如果在清单中进行的更改似乎没有生效，例如功能区按钮图标的文件名或加载项命令的文本，请尝试清除计算机上的 Office 缓存。 

#### <a name="for-windows"></a>对于 Windows：
删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。

#### <a name="for-mac"></a>对于 Mac：
删除文件夹 `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` 的内容。

#### <a name="for-ios"></a>对于 iOS：
从加载项中的 JavaScript 调用 `window.location.reload(true)` 以强制重载。也可以重新安装 Office。

## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](../develop/add-in-manifests.md)
- [旁加载 Office 加载项以供测试](sideload-office-add-ins-for-testing.md)
- [调试 Office加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
