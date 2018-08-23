---
title: 验证并排查清单问题
description: 使用这些方法验证 Office 加载项清单。
ms.date: 12/04/2017
ms.openlocfilehash: 19f7caaf1d5482972432aad3d2774d69c75cde76
ms.sourcegitcommit: 7ecc1dc24bf7488b53117d7a83ad60e952a6f7aa
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/23/2018
ms.locfileid: "19438758"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="72f8c-103">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="72f8c-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="72f8c-104">下面这些方法可用于验证和排查 Office 加载项清单问题：</span><span class="sxs-lookup"><span data-stu-id="72f8c-104">Use these methods to validate and troubleshoot issues in your Office Add-ins manifest:</span></span> 

- [<span data-ttu-id="72f8c-105">通过 Office 加载项验证程序验证清单</span><span class="sxs-lookup"><span data-stu-id="72f8c-105">Validate your manifest with the Office Add-in Validator</span></span>](#validate-your-manifest-with-the-office-add-in-validator)   
- [<span data-ttu-id="72f8c-106">根据 XML 架构验证清单</span><span class="sxs-lookup"><span data-stu-id="72f8c-106">Validate your manifest against the XML schema</span></span>](#validate-your-manifest-against-the-xml-schema)
- [<span data-ttu-id="72f8c-107">使用运行时日志记录功能调试加载项清单</span><span class="sxs-lookup"><span data-stu-id="72f8c-107">Use runtime logging to debug your add-in manifest</span></span>](#use-runtime-logging-to-debug-your-add-in-manifest)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a><span data-ttu-id="72f8c-108">通过 Office 加载项验证程序验证清单</span><span class="sxs-lookup"><span data-stu-id="72f8c-108">Validate your manifest with the Office Add-in Validator</span></span>

<span data-ttu-id="72f8c-109">为了确保描述 Office 加载项的清单文件正确完整，请使用 [Office 加载项验证程序](https://github.com/OfficeDev/office-addin-validator)验证清单。</span><span class="sxs-lookup"><span data-stu-id="72f8c-109">To help ensure that the manifest file that describes your Office Add-in is correct and complete, validate it against the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator).</span></span>

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a><span data-ttu-id="72f8c-110">使用 Office 加载项验证程序验证清单的具体步骤</span><span class="sxs-lookup"><span data-stu-id="72f8c-110">To use the Office Add-in Validator to validate your manifest</span></span>

1. <span data-ttu-id="72f8c-111">安装 [Node.js](https://nodejs.org/download/)。</span><span class="sxs-lookup"><span data-stu-id="72f8c-111">Install [Node.js](https://nodejs.org/download/).</span></span> 

2. <span data-ttu-id="72f8c-112">以管理员身份打开命令提示符/终端，并运行下面的命令，以全局安装 Office 加载项验证程序及其依赖项：</span><span class="sxs-lookup"><span data-stu-id="72f8c-112">Open a command prompt / terminal as an administrator, and install the Office Add-in Validator and its dependencies globally by using the following command:</span></span>

    ```bash
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > <span data-ttu-id="72f8c-113">如果已安装 Yo Office，请升级到最新版本，验证程序就会作为依赖项进行安装。</span><span class="sxs-lookup"><span data-stu-id="72f8c-113">If you already have Yo Office installed, upgrade to the latest version, and the validator will be installed as a dependency.</span></span>

3. <span data-ttu-id="72f8c-p101">运行下面的命令来验证清单。将 MANIFEST.XML 替换为清单 XML 文件路径。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p101">Run the following command to validate your manifest. Replace MANIFEST.XML with the path to the manifest XML file.</span></span>

    ```bash
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="72f8c-116">根据 XML 架构验证清单</span><span class="sxs-lookup"><span data-stu-id="72f8c-116">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="72f8c-117">为了有助于确保清单文件采用正确架构，请为要使用的元素添加任何命名空间。</span><span class="sxs-lookup"><span data-stu-id="72f8c-117">To help ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="72f8c-118">如果从其他示例清单中复制了元素，请仔细检查是否还**添加了相应命名空间**。</span><span class="sxs-lookup"><span data-stu-id="72f8c-118">If you copied elements from other sample manifests double check you also **include the appropiate namespaces**.</span></span> <span data-ttu-id="72f8c-119">可以根据 [XML 架构定义 (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件验证清单。</span><span class="sxs-lookup"><span data-stu-id="72f8c-119">You can validate a manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="72f8c-120">若要执行此验证，可以使用 XML 架构验证工具。</span><span class="sxs-lookup"><span data-stu-id="72f8c-120">You can use an XML schema validation tool to perform this validation.</span></span> 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="72f8c-121">使用命令行 XML 架构验证工具验证清单的具体步骤</span><span class="sxs-lookup"><span data-stu-id="72f8c-121">To use a command-line XML schema validation tool to validate your manifest</span></span>

1.  <span data-ttu-id="72f8c-122">安装 [tar](https://www.gnu.org/software/tar/) 和 [libxml](http://xmlsoft.org/FAQ.html)（如果尚未安装的话）。</span><span class="sxs-lookup"><span data-stu-id="72f8c-122">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2.  <span data-ttu-id="72f8c-p103">运行下面的命令。将 `XSD_FILE` 替换为清单 XSD 文件路径，并将 `XML_FILE` 替换为清单 XML 文件路径。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p103">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```bash
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in-manifest"></a><span data-ttu-id="72f8c-125">使用运行时日志记录功能调试加载项清单</span><span class="sxs-lookup"><span data-stu-id="72f8c-125">Use runtime logging to debug your add-in manifest</span></span>

<span data-ttu-id="72f8c-p104">可以使用运行时日志记录调试外接程序的清单。此功能可以帮助你标识并修复清单中未被 XSD 架构验证检测到的问题，例如资源 ID 间的不匹配等。运行时日志记录对于调试执行外接程序命令的外接程序尤其有用。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p104">You can use runtime logging to debug your add-in's manifest. This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs. Runtime logging is particularly  useful for debugging add-ins that implement add-in commands.</span></span>  

> [!NOTE]
> <span data-ttu-id="72f8c-129">运行时日志记录功能暂适用于 Office 2016 桌面版。</span><span class="sxs-lookup"><span data-stu-id="72f8c-129">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

### <a name="to-turn-on-runtime-logging"></a><span data-ttu-id="72f8c-130">启用运行时日志记录功能的具体步骤</span><span class="sxs-lookup"><span data-stu-id="72f8c-130">To turn on runtime logging</span></span>

> [!IMPORTANT]
> <span data-ttu-id="72f8c-p105">运行时日志记录功能影响性能。仅在需要调试加载项清单问题时，才启用此功能。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p105">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

<span data-ttu-id="72f8c-133">若要启用运行时日志记录功能，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="72f8c-133">To turn on runtime logging:</span></span>

1. <span data-ttu-id="72f8c-134">确保运行的是 Office 2016 桌面版 **16.0.7019** 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="72f8c-134">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="72f8c-135">在 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\` 下添加 `RuntimeLogging` 注册表项。</span><span class="sxs-lookup"><span data-stu-id="72f8c-135">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`.</span></span> 

3. <span data-ttu-id="72f8c-p106">将此项的默认值设置为你想要在其中写入日志的文件的完整路径。有关示例，请参阅 [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip)。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p106">Set the default value of the key to the full path of the file where you want the log to be written. For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="72f8c-138">向其中写入日志文件的目录必须已经存在，并且必须拥有对它的写入权限。</span><span class="sxs-lookup"><span data-stu-id="72f8c-138">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="72f8c-139">注册表应如下图所示。</span><span class="sxs-lookup"><span data-stu-id="72f8c-139">The following image shows what the registry should look like.</span></span> <span data-ttu-id="72f8c-140">若要禁用此功能，请从注册表中删除 `RuntimeLogging`。</span><span class="sxs-lookup"><span data-stu-id="72f8c-140">To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![包含 RuntimeLogging 注册表项的注册表编辑器屏幕截图](http://i.imgur.com/Sa9TyI6.png)


### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="72f8c-142">排查清单问题的具体步骤</span><span class="sxs-lookup"><span data-stu-id="72f8c-142">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="72f8c-143">若要使用运行时日志记录功能排查加载项的加载问题，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="72f8c-143">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="72f8c-144">[旁加载加载项](sideload-office-add-ins-for-testing.md)以供测试。</span><span class="sxs-lookup"><span data-stu-id="72f8c-144">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="72f8c-145">建议仅旁加载要测试的加载项，以最大限度地减少日志文件中的消息数。</span><span class="sxs-lookup"><span data-stu-id="72f8c-145">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="72f8c-146">如果没有任何变化，且看不到加载项（加载项对话框中没有显示），请打开日志文件。</span><span class="sxs-lookup"><span data-stu-id="72f8c-146">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="72f8c-p108">在日志文件中搜索你的外接程序 ID（已在清单中定义）。在日志文件中，此 ID 标有 `SolutionId`。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p108">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="72f8c-p109">在以下示例中，日志文件标识指向某个不存在的资源文件的控件。对于此示例，修复方法则是更正清单中的输入错误或添加丢失的资源。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p109">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![带有指定未找到的资源 ID 的条目的日志文件屏幕截图](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="72f8c-152">运行时日志记录已知问题</span><span class="sxs-lookup"><span data-stu-id="72f8c-152">Known issues with runtime logging</span></span>

<span data-ttu-id="72f8c-p110">在日志文件中看到的信息可能很混乱或其分类不正确。例如：</span><span class="sxs-lookup"><span data-stu-id="72f8c-p110">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="72f8c-155">后跟 `Unexpected Parsed manifest targeting different host` 的信息 `Medium Current host not in add-in's host list` 是错误分类为错误。</span><span class="sxs-lookup"><span data-stu-id="72f8c-155">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="72f8c-156">如果发现信息 `Unexpected Add-in is missing required manifest fields DisplayName` 且其不包含 SolutionId，则此错误极可能与你正在调试的外接程序无关。</span><span class="sxs-lookup"><span data-stu-id="72f8c-156">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="72f8c-p111">对系统而言，任何 `Monitorable` 信息都会视其为错误。有时这些信息表示清单中的问题，例如一个已跳过但未引起清单运行失败的拼写错误的元素。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p111">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="72f8c-159">清除 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="72f8c-159">Clear the Office cache</span></span>

<span data-ttu-id="72f8c-160">如果在清单中进行的更改（如功能区按钮图标的文件名或加载项命令的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="72f8c-160">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="72f8c-161">对于 Windows：</span><span class="sxs-lookup"><span data-stu-id="72f8c-161">For Windows:</span></span>
<span data-ttu-id="72f8c-162">删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。</span><span class="sxs-lookup"><span data-stu-id="72f8c-162">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="72f8c-163">对于 Mac：</span><span class="sxs-lookup"><span data-stu-id="72f8c-163">For Mac:</span></span>
<span data-ttu-id="72f8c-164">删除文件夹 `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` 的内容。</span><span class="sxs-lookup"><span data-stu-id="72f8c-164">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="72f8c-165">对于 iOS：</span><span class="sxs-lookup"><span data-stu-id="72f8c-165">For iOS:</span></span>
<span data-ttu-id="72f8c-p112">在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="72f8c-p112">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="72f8c-168">另请参阅</span><span class="sxs-lookup"><span data-stu-id="72f8c-168">See also</span></span>

- [<span data-ttu-id="72f8c-169">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="72f8c-169">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="72f8c-170">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="72f8c-170">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="72f8c-171">调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="72f8c-171">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
