---
title: 使用运行时日志记录功能调试加载项
description: 了解如何使用运行时日志记录功能调试加载项。
ms.date: 09/23/2020
localization_priority: Normal
ms.openlocfilehash: 6ba3c4cf4d94007cd0dc96480a7805f507d358c2
ms.sourcegitcommit: 42202d7e2ac24dffa77cf937f5697a1cd79ee790
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2020
ms.locfileid: "48308528"
---
# <a name="debug-your-add-in-with-runtime-logging"></a><span data-ttu-id="6a9df-103">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="6a9df-103">Debug your add-in with runtime logging</span></span>

<span data-ttu-id="6a9df-104">可以使用运行时日志记录调试加载项的清单以及多个安装错误。</span><span class="sxs-lookup"><span data-stu-id="6a9df-104">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="6a9df-105">此功能可以帮助你标识并修复清单中未被 XSD 架构验证检测到的问题，例如资源 ID 间的不匹配等。</span><span class="sxs-lookup"><span data-stu-id="6a9df-105">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="6a9df-106">运行时日志记录对于调试执行加载项命令的加载项和 Excel 自定义功能尤其有用。</span><span class="sxs-lookup"><span data-stu-id="6a9df-106">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>

> [!NOTE]
> <span data-ttu-id="6a9df-107">运行时日志记录功能当前可用于 Office 2016 或更高版本在桌面上。</span><span class="sxs-lookup"><span data-stu-id="6a9df-107">The runtime logging feature is currently available for Office 2016 or later on desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6a9df-108">运行时日志记录影响性能。</span><span class="sxs-lookup"><span data-stu-id="6a9df-108">Runtime Logging affects performance.</span></span> <span data-ttu-id="6a9df-109">请仅在需要调试外接程序清单中的问题时启用此功能。</span><span class="sxs-lookup"><span data-stu-id="6a9df-109">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

## <a name="use-runtime-logging-from-the-command-line"></a><span data-ttu-id="6a9df-110">使用命令行中的运行时日志</span><span class="sxs-lookup"><span data-stu-id="6a9df-110">Use runtime logging from the command line</span></span>

<span data-ttu-id="6a9df-111">从命令行启用运行时日志记录是最快的使用此日志记录工具的方式。</span><span class="sxs-lookup"><span data-stu-id="6a9df-111">Enabling runtime logging from the command line is the fastest way to use this logging tool.</span></span> <span data-ttu-id="6a9df-112">这些使用 npx，默认情况下，它作为 npm@5.2.0+ 的一部分提供。</span><span class="sxs-lookup"><span data-stu-id="6a9df-112">These use npx, which is provided by default as part of npm@5.2.0+.</span></span> <span data-ttu-id="6a9df-113">如果使用的是 [npm](https://www.npmjs.com/) 的早期版本，请尝试 [Windows 上的运行时日志记录](#runtime-logging-on-windows)或 [Mac](#runtime-logging-on-mac)说明，或者[安装 npx](https://www.npmjs.com/package/npx)。</span><span class="sxs-lookup"><span data-stu-id="6a9df-113">If you have an earlier version of [npm](https://www.npmjs.com/), try [Runtime logging on Windows](#runtime-logging-on-windows) or [Runtime logging on Mac](#runtime-logging-on-mac) instructions, or [install npx](https://www.npmjs.com/package/npx).</span></span>

- <span data-ttu-id="6a9df-114">要启用运行时日志记录，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6a9df-114">To enable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```

- <span data-ttu-id="6a9df-115">若要仅对特定文件启用运行时日志记录，请使用包含文件名的相同命令：</span><span class="sxs-lookup"><span data-stu-id="6a9df-115">To enable runtime logging only for a specific file, use the same command with a filename:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- <span data-ttu-id="6a9df-116">要禁用运行时日志记录，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6a9df-116">To disable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- <span data-ttu-id="6a9df-117">要显示是否启用了运行时日志记录，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6a9df-117">To display whether runtime logging is enabled:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- <span data-ttu-id="6a9df-118">要在命令行内显示有关运行时日志记录的帮助，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6a9df-118">To display help within the command line for runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## <a name="runtime-logging-on-windows"></a><span data-ttu-id="6a9df-119">Windows 上的运行时日志记录</span><span class="sxs-lookup"><span data-stu-id="6a9df-119">Runtime logging on Windows</span></span>

1. <span data-ttu-id="6a9df-120">确保运行的是 Office 2016 桌面版 **16.0.7019** 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="6a9df-120">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span>

2. <span data-ttu-id="6a9df-121">在 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\` 下添加 `RuntimeLogging` 注册表项。</span><span class="sxs-lookup"><span data-stu-id="6a9df-121">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a9df-122">如果 `Developer` (文件夹中尚不存在) 的项 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\` ，请完成以下步骤以创建它。</span><span class="sxs-lookup"><span data-stu-id="6a9df-122">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it.</span></span>
    >
    > 1. <span data-ttu-id="6a9df-123">右键单击 **WEF** 密钥（文件夹），然后选择**新建** > **密钥**。</span><span class="sxs-lookup"><span data-stu-id="6a9df-123">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 1. <span data-ttu-id="6a9df-124">将新密钥命名为 **Developer**。</span><span class="sxs-lookup"><span data-stu-id="6a9df-124">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="6a9df-125">将 **RuntimeLogging** 项的默认值设置为你想要在其中写入日志的文件的完整路径。</span><span class="sxs-lookup"><span data-stu-id="6a9df-125">Set the default value of the **RuntimeLogging** key to the full path of the file where you want the log to be written.</span></span> <span data-ttu-id="6a9df-126">有关示例，请参阅 [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip)。</span><span class="sxs-lookup"><span data-stu-id="6a9df-126">For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a9df-127">向其中写入日志文件的目录必须已经存在，并且必须拥有对它的写入权限。</span><span class="sxs-lookup"><span data-stu-id="6a9df-127">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span>

<span data-ttu-id="6a9df-p105">注册表应如下图所示。 若要禁用此功能，请从注册表中删除 `RuntimeLogging`。</span><span class="sxs-lookup"><span data-stu-id="6a9df-p105">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span>

![包含 RuntimeLogging 注册表项的注册表编辑器屏幕截图](http://i.imgur.com/Sa9TyI6.png)

## <a name="runtime-logging-on-mac"></a><span data-ttu-id="6a9df-131">Mac 上的运行时日志记录</span><span class="sxs-lookup"><span data-stu-id="6a9df-131">Runtime logging on Mac</span></span>

1. <span data-ttu-id="6a9df-132">确保运行的是 Office 2016 桌面版 **16.27** (19071500) 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="6a9df-132">Make sure that you are running Office 2016 desktop build **16.27** (19071500) or later.</span></span>

2. <span data-ttu-id="6a9df-133">打开**终端**并使用 `defaults` 命令设置运行时日志记录首选项：</span><span class="sxs-lookup"><span data-stu-id="6a9df-133">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>

    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="6a9df-134">`<bundle id>` 确定了主机要对哪些运行时日志记录。</span><span class="sxs-lookup"><span data-stu-id="6a9df-134">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="6a9df-135">`<file_name>` 是要将日志写入的文本文件的名称。</span><span class="sxs-lookup"><span data-stu-id="6a9df-135">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="6a9df-136">设置 `<bundle id>` 为以下值之一，以便为相应的应用程序启用运行时日志记录：</span><span class="sxs-lookup"><span data-stu-id="6a9df-136">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding application:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="6a9df-137">以下示例会为 Word 启用运行时日志记录，然后打开日志文件：</span><span class="sxs-lookup"><span data-stu-id="6a9df-137">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE]
> <span data-ttu-id="6a9df-138">运行 `defaults` 命令来启用运行时日志记录后，需要重启 Office。</span><span class="sxs-lookup"><span data-stu-id="6a9df-138">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="6a9df-139">要关闭运行时日志记录，请使用 `defaults delete` 命令：</span><span class="sxs-lookup"><span data-stu-id="6a9df-139">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="6a9df-140">以下示例将为 Word 关闭运行时日志记录：</span><span class="sxs-lookup"><span data-stu-id="6a9df-140">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## <a name="use-runtime-logging-to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="6a9df-141">使用运行时日志记录功能排查清单问题</span><span class="sxs-lookup"><span data-stu-id="6a9df-141">Use runtime logging to troubleshoot issues with your manifest</span></span>

<span data-ttu-id="6a9df-142">若要使用运行时日志记录功能排查加载项的加载问题，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6a9df-142">To use runtime logging to troubleshoot issues loading an add-in:</span></span>

1. <span data-ttu-id="6a9df-143">[旁加载加载项](sideload-office-add-ins-for-testing.md)以供测试。</span><span class="sxs-lookup"><span data-stu-id="6a9df-143">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6a9df-144">建议仅旁加载要测试的加载项，以最大限度地减少日志文件中的消息数。</span><span class="sxs-lookup"><span data-stu-id="6a9df-144">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="6a9df-145">如果没有任何变化，且看不到加载项（加载项对话框中没有显示），请打开日志文件。</span><span class="sxs-lookup"><span data-stu-id="6a9df-145">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="6a9df-p107">在日志文件中搜索你的外接程序 ID（已在清单中定义）。在日志文件中，此 ID 标有 `SolutionId`。</span><span class="sxs-lookup"><span data-stu-id="6a9df-p107">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span>

<span data-ttu-id="6a9df-p108">在以下示例中，日志文件标识指向某个不存在的资源文件的控件。对于此示例，修复方法则是更正清单中的输入错误或添加丢失的资源。</span><span class="sxs-lookup"><span data-stu-id="6a9df-p108">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![带有指定未找到的资源 ID 的条目的日志文件屏幕截图](http://i.imgur.com/f8bouLA.png)

## <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="6a9df-151">运行时日志记录已知问题</span><span class="sxs-lookup"><span data-stu-id="6a9df-151">Known issues with runtime logging</span></span>

<span data-ttu-id="6a9df-p109">在日志文件中看到的信息可能很混乱或其分类不正确。例如：</span><span class="sxs-lookup"><span data-stu-id="6a9df-p109">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="6a9df-154">后跟 `Unexpected Parsed manifest targeting different host` 的信息 `Medium Current host not in add-in's host list` 是错误分类为错误。</span><span class="sxs-lookup"><span data-stu-id="6a9df-154">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="6a9df-155">如果发现信息 `Unexpected Add-in is missing required manifest fields    DisplayName` 且其不包含 SolutionId，则此错误极可能与你正在调试的外接程序无关。</span><span class="sxs-lookup"><span data-stu-id="6a9df-155">If you see the message `Unexpected Add-in is missing required manifest fields    DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span>

- <span data-ttu-id="6a9df-p110">对系统而言，任何 `Monitorable` 信息都会视其为错误。有时这些信息表示清单中的问题，例如一个已跳过但未引起清单运行失败的拼写错误的元素。</span><span class="sxs-lookup"><span data-stu-id="6a9df-p110">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span>

## <a name="see-also"></a><span data-ttu-id="6a9df-158">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6a9df-158">See also</span></span>

- [<span data-ttu-id="6a9df-159">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="6a9df-159">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="6a9df-160">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="6a9df-160">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="6a9df-161">清除 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="6a9df-161">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="6a9df-162">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="6a9df-162">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="6a9df-163">调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="6a9df-163">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)