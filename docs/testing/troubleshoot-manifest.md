---
title: 验证并排查清单问题
description: 使用这些方法验证 Office 加载项清单。
ms.date: 09/18/2019
localization_priority: Priority
ms.openlocfilehash: c320c05b944bba9e24a4d3c0e5ef514ac13cc3c6
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035334"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="cbadb-103">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="cbadb-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="cbadb-104">你可能需要验证加载项的清单文件，以确保其正确且完整。</span><span class="sxs-lookup"><span data-stu-id="cbadb-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="cbadb-105">当你尝试旁加载加载项时，验证还可以识别导致错误“你的加载项清单无效”的问题。</span><span class="sxs-lookup"><span data-stu-id="cbadb-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="cbadb-106">本文介绍了验证清单文件和解决加载项问题的多种方法。</span><span class="sxs-lookup"><span data-stu-id="cbadb-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="cbadb-107">使用适用于 Office 加载项的 Yeoman 生成器来验证清单</span><span class="sxs-lookup"><span data-stu-id="cbadb-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="cbadb-108">如果你使用了[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则也可以使用它来验证项目的清单文件。</span><span class="sxs-lookup"><span data-stu-id="cbadb-108">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="cbadb-109">在项目的根目录中运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="cbadb-109">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![动画 gif 显示 Yo Office 验证程序（在命令行处运行并生成显示“验证已通过”的结果）。](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="cbadb-111">若要访问此功能，必须使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)版本 1.1.17 或更高版本创建加载项项目。</span><span class="sxs-lookup"><span data-stu-id="cbadb-111">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="cbadb-112">使用 office-addin-manifest 验证清单</span><span class="sxs-lookup"><span data-stu-id="cbadb-112">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="cbadb-113">如果你未使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则可以使用 [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest)。</span><span class="sxs-lookup"><span data-stu-id="cbadb-113">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="cbadb-114">安装 [Node.js](https://nodejs.org/download/)。</span><span class="sxs-lookup"><span data-stu-id="cbadb-114">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="cbadb-115">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="cbadb-115">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="cbadb-116">将 `MANIFEST_FILE` 替换为清单文件的名称。</span><span class="sxs-lookup"><span data-stu-id="cbadb-116">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="cbadb-117">如果运行此命令导致错误消息“命令语法无效。”</span><span class="sxs-lookup"><span data-stu-id="cbadb-117">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="cbadb-118">（因为 `validate` 命令无法识别），运行以下命令验证清单（用清单文件的名称替换 `MANIFEST_FILE`）：</span><span class="sxs-lookup"><span data-stu-id="cbadb-118">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="cbadb-119">根据 XML 架构验证清单</span><span class="sxs-lookup"><span data-stu-id="cbadb-119">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="cbadb-120">可以根据 [XML 架构定义 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件来验证清单文件。</span><span class="sxs-lookup"><span data-stu-id="cbadb-120">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="cbadb-121">这将有助于确保清单文件采用正确架构（包括所使用的元素的所有命名空间）。</span><span class="sxs-lookup"><span data-stu-id="cbadb-121">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="cbadb-122">如果从其他示例清单中复制了元素，请仔细检查是否还**添加了相应命名空间**。</span><span class="sxs-lookup"><span data-stu-id="cbadb-122">If you copied elements from other sample manifests double check you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="cbadb-123">若要执行此验证，可以使用 XML 架构验证工具。</span><span class="sxs-lookup"><span data-stu-id="cbadb-123">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="cbadb-124">使用命令行 XML 架构验证工具验证清单的具体步骤</span><span class="sxs-lookup"><span data-stu-id="cbadb-124">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="cbadb-125">安装 [tar](https://www.gnu.org/software/tar/) 和 [libxml](http://xmlsoft.org/FAQ.html)（如果尚未安装的话）。</span><span class="sxs-lookup"><span data-stu-id="cbadb-125">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="cbadb-p106">运行下面的命令。将 `XSD_FILE` 替换为清单 XSD 文件路径，并将 `XML_FILE` 替换为清单 XML 文件路径。</span><span class="sxs-lookup"><span data-stu-id="cbadb-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="cbadb-128">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="cbadb-128">Use runtime logging to debug your add-in</span></span>

<span data-ttu-id="cbadb-129">可以使用运行时日志记录调试加载项的清单以及多个安装错误。</span><span class="sxs-lookup"><span data-stu-id="cbadb-129">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="cbadb-130">此功能可以帮助你标识并修复清单中未被 XSD 架构验证检测到的问题，例如资源 ID 间的不匹配等。</span><span class="sxs-lookup"><span data-stu-id="cbadb-130">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="cbadb-131">运行时日志记录对于调试执行加载项命令的加载项和 Excel 自定义功能尤其有用。</span><span class="sxs-lookup"><span data-stu-id="cbadb-131">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="cbadb-132">运行时日志记录功能暂适用于 Office 2016 桌面版。</span><span class="sxs-lookup"><span data-stu-id="cbadb-132">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cbadb-133">运行时日志记录影响性能。</span><span class="sxs-lookup"><span data-stu-id="cbadb-133">Runtime Logging affects performance.</span></span> <span data-ttu-id="cbadb-134">请仅在需要调试外接程序清单中的问题时启用此功能。</span><span class="sxs-lookup"><span data-stu-id="cbadb-134">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

### <a name="runtime-logging-on-windows"></a><span data-ttu-id="cbadb-135">Windows 上的运行时日志记录</span><span class="sxs-lookup"><span data-stu-id="cbadb-135">Runtime logging on Windows</span></span>

1. <span data-ttu-id="cbadb-136">确保运行的是 Office 2016 桌面版 **16.0.7019** 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="cbadb-136">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="cbadb-137">在 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` 下添加 `RuntimeLogging` 注册表项。</span><span class="sxs-lookup"><span data-stu-id="cbadb-137">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="cbadb-138">如果 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\` 下尚不存在 `Developer` 密钥（文件夹），请完成以下步骤以创建它：</span><span class="sxs-lookup"><span data-stu-id="cbadb-138">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="cbadb-139">右键单击 **WEF** 密钥（文件夹），然后选择**新建** > **密钥**。</span><span class="sxs-lookup"><span data-stu-id="cbadb-139">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="cbadb-140">将新密钥命名为 **Developer**。</span><span class="sxs-lookup"><span data-stu-id="cbadb-140">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="cbadb-p109">将此项的默认值设置为你想要在其中写入日志的文件的完整路径。有关示例，请参阅 [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip)。</span><span class="sxs-lookup"><span data-stu-id="cbadb-p109">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="cbadb-143">向其中写入日志文件的目录必须已经存在，并且必须拥有对它的写入权限。</span><span class="sxs-lookup"><span data-stu-id="cbadb-143">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="cbadb-p110">注册表应如下图所示。 若要禁用此功能，请从注册表中删除 `RuntimeLogging`。</span><span class="sxs-lookup"><span data-stu-id="cbadb-p110">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![包含 RuntimeLogging 注册表项的注册表编辑器屏幕截图](http://i.imgur.com/Sa9TyI6.png)

### <a name="runtime-logging-on-mac"></a><span data-ttu-id="cbadb-147">Mac 上的运行时日志记录</span><span class="sxs-lookup"><span data-stu-id="cbadb-147">Runtime logging on Mac</span></span>

1. <span data-ttu-id="cbadb-148">确保运行的是 Office 2016 桌面版 **16.27** (19071500) 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="cbadb-148">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span>

2. <span data-ttu-id="cbadb-149">打开**终端**并使用 `defaults` 命令设置运行时日志记录首选项：</span><span class="sxs-lookup"><span data-stu-id="cbadb-149">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="cbadb-150">`<bundle id>` 确定了主机要对哪些运行时日志记录。</span><span class="sxs-lookup"><span data-stu-id="cbadb-150">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="cbadb-151">`<file_name>` 是要将日志写入的文本文件的名称。</span><span class="sxs-lookup"><span data-stu-id="cbadb-151">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="cbadb-152">将 `<bundle id>` 设置为下述值之一，从而为相应主机启用运行时日志记录：</span><span class="sxs-lookup"><span data-stu-id="cbadb-152">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding host:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="cbadb-153">以下示例会为 Word 启用运行时日志记录，然后打开日志文件：</span><span class="sxs-lookup"><span data-stu-id="cbadb-153">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> <span data-ttu-id="cbadb-154">运行 `defaults` 命令来启用运行时日志记录后，需要重启 Office。</span><span class="sxs-lookup"><span data-stu-id="cbadb-154">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="cbadb-155">要关闭运行时日志记录，请使用 `defaults delete` 命令：</span><span class="sxs-lookup"><span data-stu-id="cbadb-155">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="cbadb-156">以下示例将为 Word 关闭运行时日志记录：</span><span class="sxs-lookup"><span data-stu-id="cbadb-156">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="cbadb-157">排查清单问题的具体步骤</span><span class="sxs-lookup"><span data-stu-id="cbadb-157">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="cbadb-158">若要使用运行时日志记录功能排查加载项的加载问题，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="cbadb-158">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="cbadb-159">[旁加载加载项](sideload-office-add-ins-for-testing.md)以供测试。</span><span class="sxs-lookup"><span data-stu-id="cbadb-159">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="cbadb-160">建议仅旁加载要测试的加载项，以最大限度地减少日志文件中的消息数。</span><span class="sxs-lookup"><span data-stu-id="cbadb-160">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="cbadb-161">如果没有任何变化，且看不到加载项（加载项对话框中没有显示），请打开日志文件。</span><span class="sxs-lookup"><span data-stu-id="cbadb-161">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="cbadb-p112">在日志文件中搜索你的外接程序 ID（已在清单中定义）。在日志文件中，此 ID 标有 `SolutionId`。</span><span class="sxs-lookup"><span data-stu-id="cbadb-p112">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="cbadb-p113">在以下示例中，日志文件标识指向某个不存在的资源文件的控件。对于此示例，修复方法则是更正清单中的输入错误或添加丢失的资源。</span><span class="sxs-lookup"><span data-stu-id="cbadb-p113">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![带有指定未找到的资源 ID 的条目的日志文件屏幕截图](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="cbadb-167">运行时日志记录已知问题</span><span class="sxs-lookup"><span data-stu-id="cbadb-167">Known issues with runtime logging</span></span>

<span data-ttu-id="cbadb-p114">在日志文件中看到的信息可能很混乱或其分类不正确。例如：</span><span class="sxs-lookup"><span data-stu-id="cbadb-p114">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="cbadb-170">后跟 `Unexpected Parsed manifest targeting different host` 的信息 `Medium Current host not in add-in's host list` 是错误分类为错误。</span><span class="sxs-lookup"><span data-stu-id="cbadb-170">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="cbadb-171">如果发现信息 `Unexpected Add-in is missing required manifest fields DisplayName` 且其不包含 SolutionId，则此错误极可能与你正在调试的外接程序无关。</span><span class="sxs-lookup"><span data-stu-id="cbadb-171">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="cbadb-p115">对系统而言，任何 `Monitorable` 信息都会视其为错误。有时这些信息表示清单中的问题，例如一个已跳过但未引起清单运行失败的拼写错误的元素。</span><span class="sxs-lookup"><span data-stu-id="cbadb-p115">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="cbadb-174">清除 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="cbadb-174">Clear the Office cache</span></span>

<span data-ttu-id="cbadb-175">如果在清单中进行的更改（如功能区按钮图标的文件名或加载项命令的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="cbadb-175">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="cbadb-176">对于 Windows：</span><span class="sxs-lookup"><span data-stu-id="cbadb-176">For Windows:</span></span>
<span data-ttu-id="cbadb-177">删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。</span><span class="sxs-lookup"><span data-stu-id="cbadb-177">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="cbadb-178">对于 Mac：</span><span class="sxs-lookup"><span data-stu-id="cbadb-178">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="cbadb-179">对于 iOS：</span><span class="sxs-lookup"><span data-stu-id="cbadb-179">For iOS:</span></span>
<span data-ttu-id="cbadb-p116">在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="cbadb-p116">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="cbadb-182">另请参阅</span><span class="sxs-lookup"><span data-stu-id="cbadb-182">See also</span></span>

- [<span data-ttu-id="cbadb-183">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="cbadb-183">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="cbadb-184">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="cbadb-184">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="cbadb-185">调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="cbadb-185">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
