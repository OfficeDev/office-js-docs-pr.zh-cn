---
title: Office 加载项的资源限制和性能优化
description: ''
ms.date: 09/09/2019
localization_priority: Normal
ms.openlocfilehash: 332ea72e9c96ff7a9b61a4fb0249284ca44079ac
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42162788"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Office 加载项的资源限制和性能优化

为了向用户提供最佳体验，请确保 Office 加载项不超过 CPU 内核和内存使用、可靠性以及计算正则表达式的响应时间（对于 Outlook 加载项）方面的特定限制。这些运行时资源使用限制仅适用于在 Windows 和 OS X 的 Office 客户端上运行的加载项，而不适用于移动应用或浏览器上的加载项。

此外，还可以在加载项设计和实现中优化资源使用，从而优化加载项在台式机和移动设备上的性能。

## <a name="resource-usage-limits-for-add-ins"></a>加载项的资源使用限制

运行时资源使用限制适用于所有类型的 Office 加载项。这些限制有助于确保为用户提供优质性能，并减轻拒绝服务攻击。请务必使用一系列可能数据在目标主机应用上测试 Office 加载项，并根据以下运行时使用限制衡量性能：

- **CPU 内核使用** - 单个 CPU 内核使用阈值为 90%，默认每 5 秒监测三次。

   对于主机丰富客户端，检查 CPU 内核使用的默认时间间隔为 5 秒。如果主机客户端检测到加载项的 CPU 内核使用超出阈值，便会显示消息，询问用户是否要继续运行加载项。如果用户选择继续运行，主机客户端在此编辑会话期间不会再次询问用户。如果用户运行占用大量 CPU 的加载项，建议管理员使用“AlertInterval”**** 注册表项增加阈值，以减少此类警告消息的显示。

- **内存使用** - 默认内存使用阈值，根据设备的可用物理内存动态确定。

   默认情况下，当主机丰富客户端检测到某个设备上的物理内存使用率超出可用内存的 80% 时，客户端将开始监视该外接程序的内存使用率（监视内容和任务窗格外接程序的文档级别，监视 Outlook 外接程序的邮箱级别）。在默认为 5 秒的间隔时间，客户端警告用户在文档或邮箱级别中的一组外接程序的物理内存使用率是否超出 50%。该内存使用率限制使用物理内存（而非虚拟内存）来确保使用了有限 RAM 的设备的性能表现，如平板电脑。管理员可以通过使用 **MemoryAlertThreshold** Windows 注册表项作为全局设置将动态设置替代为显式限制，或通过使用 **AlertInterval** 密钥作为全局设置来调整警报间隔。

- **故障容忍度** - 外接程序的默认限制为 4 次故障。

   管理员可以通过使用 **RestartManagerRetryLimit** 注册表项来调整故障阈值。

- **应用程序阻塞** - 外接程序持续无响应时间阈值为 5 秒。

   这会影响外接程序和主机应用程序的用户体验。发生此种情况时，主机应用程序将自动重新启动文档或邮箱的所有活动外接程序（如果适用），并警告用户哪个外接程序没有响应。外接程序在执行长时间运行的任务但未定期生成处理时，可能达到该阈值。有技术可避免出现该阻塞。管理员无法替换此阈值。

### <a name="outlook-add-ins"></a>Outlook 外接程序

如果任何 Outlook 外接程序超过上述 CPU 内核或内存使用率阈值，或者故障容忍度限制，则 Outlook 会禁用该外接程序。Exchange 管理中心会显示应用程序的禁用状态。

> [!NOTE]
> 尽管只有 Outlook 丰富客户端（而不是 Outlook 网页版或移动设备）监视资源使用，如果丰富客户端禁用 Outlook 加载项，加载项也禁用于Outlook 网页版和移动设备。

除了 CPU 内核、内存和可靠性规则之外，Outlook 加载项还应在激活后遵循以下规则：

- **正则表达式响应时间** - Outlook 计算 Outlook 外接程序清单中的所有正则表达式的默认阈值为 1,000 毫秒。超过该阈值会导致 Outlook 稍后重新尝试计算。

    通过使用 Windows 注册表中的组策略或应用程序特定设置，管理员可以在 **OutlookActivationAlertThreshold** 设置中调整此 1,000 毫秒的默认阈值。

- **正则表达式重新计算** - Outlook 重新计算清单中的所有正则表达式的默认限制为三次。如果三次计算均因超过适用阈值（默认值为 1,000 毫秒或 **OutlookActivationAlertThreshold** 指定的值，如果 Windows 注册表中存在该设置）而失败，则 Outlook 将禁用该 Outlook 加载项。Exchange 管理中心会显示禁用状态，该加载项被禁止在 Outlook 富客户端、Outlook 网页版或移动设备中使用。

    通过使用 Windows 注册表中的组策略或特定于应用程序的设置，管理员可以调整该时间数，以在 **OutlookActivationManagerRetryLimit** 设置中重试该评估。

### <a name="task-pane-and-content-add-ins"></a>任务窗格和内容外接程序

如果任何内容或任务窗格外接程序超过上述 CPU 内核或内存使用率阈值，或者故障容忍度限制，则相应的主机应用程序会向用户显示一条警告。此时，用户可以执行下列操作之一：

- 重新启动外接程序。
- 取消有关超出该阈值的后续警报。理想的情况是，用户应当从文档中删除该外接程序；继续使用该外接程序可能会遇到更多性能和稳定性问题。  

## <a name="verifying-resource-usage-issues-in-the-telemetry-log"></a>验证遥测日志中的资源使用率问题

Office 提供了遥测日志，以保留本地计算机上运行的 Office 解决方案的某些事件（加载、打开、关闭和错误）的记录，包括 Office 外接程序中的资源使用率问题。如果您已设置遥测日志，则可以使用 Excel 在您的本地驱动器中的以下默认位置打开遥测日志：

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

对于遥测日志跟踪的外接程序的每个事件，都有事件的发生日期/时间、事件 ID、严重性、事件的简短描述性标题、外接程序的友好名称和唯一 ID，以及记录事件的应用程序。可刷新遥测日志以查看当前跟踪的事件。下表显示了之前在遥测日志中跟踪的 Outlook 外接程序的示例。 

|**日期/时间**|**事件 ID**|**严重性**|**标题**|**文件**|**ID**|**应用程序**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|10/8/2012 5:57:10 PM|7||外接程序清单已成功下载|Who's Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|10/8/2012 5:57:01 PM|7||外接程序清单已成功下载|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

下表列出了遥测日志通常跟踪的 Office 外接程序的事件。

|**事件 ID**|**标题**|**严重性**|**说明**|
|:-----|:-----|:-----|:-----|
|7|外接程序清单已成功下载||主机应用程序已成功加载和读取 Office 外接程序的清单。|
|8|外接程序清单未下载|关键|主机应用无法从 SharePoint 目录、企业目录或 AppSource 加载 Office 加载项的清单文件。|
|9|无法分析外接程序标记|关键|主机应用程序已加载 Office 外接程序清单，但不能读取该应用程序的 HTML 标记。|
|10|外接程序占用了太多 CPU|关键|在有限的时间内，Office 外接程序使用了超过 90% 的 CPU 资源。|
|15|由于字符串搜索超时，外接程序已被禁用||Outlook 外接程序搜索主题行和电子邮件中的信息以确定是否应通过使用正则表达式将其显示。Outlook 禁用“**文件**”列中列出的 Outlook 外接程序，因为在尝试匹配正则表达式期间其会反复超时。|
|18|外接程序已成功关闭||主机应用程序能够成功关闭 Office 外接程序。|
|19|外接程序遇到运行时错误|关键|Office 外接程序遇到一个导致它失败的问题。有关详细信息，请使用遇到错误的计算机上的 Windows 事件查看器查看“**Microsoft Office 通知**”日志。|
|20|外接程序未能验证许可|关键|无法验证 Office 外接程序的许可信息，且其可能已过期。有关详细信息，请使用遇到错误的计算机上的 Windows 事件查看器查看“**Microsoft Office 通知**”日志。|

有关详细信息，请参阅[部署遥测仪表板](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15))和[使用遥测日志排查 Office 文件和自定义解决方案](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)。


## <a name="design-and-implementation-techniques"></a>设计和实现技术

尽管 CPU 和内存使用率的资源限制、故障容忍度以及 UI 无响应仅适用于在富客户端上运行的 Office 外接程序，但如果您希望外接程序在所有支持性客户端和设备上的性能都令人满意，优化这些资源和电池的使用情况仍然是头等大事。如果您的外接程序要执行长时间运行的操作或处理大型数据集，则优化尤为重要。以下列表提供了一些技术建议，可将大量占用 CPU 或处理大量数据的操作分解成较小的块，以便外接程序能够避免过量的资源消耗且主机应用程序可以保持响应能力：

- 在外接程序需要从无限制的数据集中读取大量数据的情况下，您可以在从表格中读取数据时应用分页，或者减小每次短暂读取操作中的数据大小，而不是试图在一次操作中完成全部读取。 

   For a JavaScript and jQuery code sample that shows breaking up a potentially long-running and CPU-intensive series of inputting and outputting operations on unbounded data, see [How can I give control back (briefly) to the browser during intensive JavaScript processing?](https://stackoverflow.com/questions/210821/how-can-i-give-control-back-briefly-to-the-browser-during-intensive-javascript). This example uses the [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) method of the global object to limit the duration of input and output. It also handles the data in defined chunks instead of randomly unbounded data.

- 如果外接程序使用大量占用 CPU 的算法来处理大量数据，则您可以使用 Web Worker 在后台执行长时间运行的任务，同时在前台运行单独的脚本，例如在用户界面中显示进度。Web Worker 不会阻止用户活动并允许 HTML 页面保持响应能力。有关 Web Worker 的示例，请参阅 [Web Worker 的基本信息](https://www.html5rocks.com/en/tutorials/workers/basics/)。有关 Web Worker API 的详细信息，请参阅 [Web Worker](https://developer.mozilla.org/docs/Web/API/Web_Workers_API)。

- 如果外接程序使用大量占用 CPU 的算法，但您可以将数据输入或输出划分成较小的集合，则可以考虑创建一个 Web 服务，将数据传递给该 Web 服务以减轻 CPU 负担，然后等待异步回调。

- 针对预期的最大数据量测试加载项，并限制加载项处理的数据量不得超过此限制。


## <a name="see-also"></a>另请参阅

- [Office 加载项的隐私和安全](../concepts/privacy-and-security.md)
- [Outlook 外接程序的激活限制和 JavaScript API](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [使用 Excel JavaScript API 优化性能](../excel/performance.md)
