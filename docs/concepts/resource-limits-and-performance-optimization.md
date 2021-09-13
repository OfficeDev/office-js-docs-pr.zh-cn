---
title: Office 加载项的资源限制和性能优化
description: 了解加载项平台的资源Office，包括 CPU 和内存。
ms.date: 08/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 03050c325ffd7b67c8c7eaf5047215fce6ac70dd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149210"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Office 加载项的资源限制和性能优化

为了向用户提供最佳体验，请确保 Office 加载项不超过 CPU 内核和内存使用、可靠性以及计算正则表达式的响应时间（对于 Outlook 加载项）方面的特定限制。这些运行时资源使用限制仅适用于在 Windows 和 OS X 的 Office 客户端上运行的加载项，而不适用于移动应用或浏览器上的加载项。

此外，还可以在加载项设计和实现中优化资源使用，从而优化加载项在台式机和移动设备上的性能。

## <a name="resource-usage-limits-for-add-ins"></a>加载项的资源使用限制

运行时资源使用限制适用于所有类型的Office外接程序。这些限制有助于确保用户性能并减少拒绝服务攻击。 请务必使用一系列Office测试目标 Office 应用程序上的加载项，并针对以下运行时使用限制测量其性能。

- **CPU 内核使用** - 单个 CPU 内核使用阈值为 90%，默认每 5 秒监测三次。

   客户端检查 CPU 内核Office的默认间隔是每 5 秒。 如果Office客户端检测到外接程序的 CPU 内核使用率超过阈值，则会显示一条消息，询问用户是否要继续运行外接程序。 如果用户选择继续，Office客户端不会在编辑会话期间再次询问用户。 如果用户运行占用大量 CPU 的加载项，建议管理员使用“AlertInterval”**** 注册表项增加阈值，以减少此类警告消息的显示。

- **内存使用** - 默认内存使用阈值，根据设备的可用物理内存动态确定。

   默认情况下，当 Office 客户端检测到设备上的物理内存使用率超过可用内存的 80% 时，客户端将开始监视外接程序的内存使用率、内容和任务窗格外接程序的文档级别以及 Outlook 外接程序的邮箱级别。在默认间隔 5 秒时，如果文档或邮箱级别的一组外接程序的物理内存使用率超过 50%，客户端将警告用户。 此内存使用率限制使用物理内存而非虚拟内存来确保具有有限 RAM 的设备（如平板电脑）保持良好性能。 管理员可以使用 **MemoryAlertThreshold** Windows 注册表项作为全局设置来覆盖此动态设置，将 **AlertInterval** 项用作全局设置来调整警报间隔。

- **故障容忍度** - 外接程序的默认限制为 4 次故障。

   管理员可以通过使用 **RestartManagerRetryLimit** 注册表项来调整故障阈值。

- **应用程序阻塞** - 外接程序持续无响应时间阈值为 5 秒。

   这会影响加载项和加载项应用程序Office体验。 发生这种情况时，Office 应用程序会自动重新启动文档或邮箱 (（如果适用) ）的所有活动外接程序，并警告用户哪个外接程序没有响应。 外接程序在执行长时间运行的任务但未定期生成处理时，可能达到该阈值。 有技术可避免出现该阻塞。 管理员无法替换此阈值。

### <a name="outlook-add-ins"></a>Outlook 外接程序

如果任何 Outlook 外接程序超过上述 CPU 内核或内存使用率阈值，或者故障容忍度限制，则 Outlook 会禁用该外接程序。Exchange 管理中心会显示应用程序的禁用状态。

> [!NOTE]
> 尽管只有 Outlook 丰富客户端（而不是 Outlook 网页版或移动设备）监视资源使用，如果丰富客户端禁用 Outlook 加载项，加载项也禁用于Outlook 网页版和移动设备。

除了 CPU 内核、内存和可靠性规则，Outlook加载项还应遵守以下激活规则。

- **正则表达式响应时间** - Outlook 计算 Outlook 外接程序清单中的所有正则表达式的默认阈值为 1,000 毫秒。超过该阈值会导致 Outlook 稍后重新尝试计算。

    通过使用组策略或 Windows 注册表中特定于应用程序的设置，管理员可以在 **OutlookActivationAlertThreshold** 设置中调整此 1,000 毫秒的默认阈值。

- **正则表达式重新计算** - Outlook 重新计算清单中的所有正则表达式的默认限制为三次。 如果计算因超过适用的阈值 (（默认值为 1，000 毫秒或 **OutlookActivationAlertThreshold** 指定的值）而全部三次失败，如果该设置存在于 Windows 注册表) 中，Outlook 将禁用 Outlook 外接程序。 Exchange管理中心显示禁用状态，并且外接程序被禁用，以在 Outlook 富客户端、Outlook 网页版移动设备中使用。

    通过使用组策略或 Windows 注册表中特定于应用程序的设置，管理员可以在 **OutlookActivationManagerRetryLimit** 设置中调整此重试计算的次数。

### <a name="excel-add-ins"></a>Excel 加载项

如果您要构建一个Excel，请注意与工作簿交互时的以下大小限制。

- Excel 网页版将请求和响应的有效负载大小限制为 5MB。 如果超过该限制，将引发 `RichAPI.Error`。
- 一个范围限制为五百万个单元格，用于获取操作。

如果你希望用户输入超出这些限制，请确保在调用 之前检查数据 `context.sync()` 。 根据需要将操作拆分为多个较小的部分。 请务必调用 `context.sync()` 每个子操作，以避免这些操作再次被批处理在一起。

这些限制通常由较大范围超出。 您的外接程序可能可以使用 [RangeAreas](/javascript/api/excel/excel.rangeareas) 以战略方式更新更大范围内的单元格。 有关使用的信息，请参阅在加载项中同时处理Excel `RangeAreas` [区域](../excel/excel-add-ins-multiple-ranges.md)。有关优化有效负载大小的其他Excel，请参阅有效负载[大小限制最佳实践](../excel/performance.md#payload-size-limit-best-practices)。

### <a name="task-pane-and-content-add-ins"></a>任务窗格和内容外接程序

如果任何内容或任务窗格外接程序超过上述 CPU 内核或内存使用率阈值，或者故障容忍度限制，则相应的 Office 应用程序会向用户显示一条警告。 此时，用户可以执行下列操作之一：

- 重新启动外接程序。
- 取消有关超出该阈值的后续警报。理想的情况是，用户应当从文档中删除该外接程序；继续使用该外接程序可能会遇到更多性能和稳定性问题。  

## <a name="verify-resource-usage-issues-in-the-telemetry-log"></a>验证遥测日志中的资源使用率问题

Office 提供了遥测日志，以保留本地计算机上运行的 Office 解决方案的某些事件（加载、打开、关闭和错误）的记录，包括 Office 外接程序中的资源使用率问题。 如果已设置遥测日志，可以使用Excel本地驱动器上的以下默认位置打开遥测日志。

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

对于遥测日志跟踪的外接程序的每个事件，都有事件的发生日期/时间、事件 ID、严重性、事件的简短描述性标题、外接程序的友好名称和唯一 ID，以及记录事件的应用程序。可刷新遥测日志以查看当前跟踪的事件。下表显示了之前在遥测日志中跟踪的 Outlook 外接程序的示例。

|**日期/时间**|**事件 ID**|**严重性**|**标题**|**文件**|**ID**|**应用程序**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|10/8/2012 5:57:10 PM|7 ||外接程序清单已成功下载|Who's Who|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|10/8/2012 5:57:01 PM|7 ||外接程序清单已成功下载|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

下表列出了遥测日志通常跟踪的 Office 外接程序的事件。

|**事件 ID**|**标题**|**严重性**|**说明**|
|:-----|:-----|:-----|:-----|
|7 |外接程序清单已成功下载||应用程序已成功Office和读取加载项Office清单。|
|8 |外接程序清单未下载|关键|Office应用程序无法从 Office 目录、公司目录或 AppSource SharePoint加载项的清单文件。|
|9 |无法分析外接程序标记|关键|加载项Office加载Office加载项清单，但无法读取应用程序的 HTML 标记。|
|10 |外接程序占用了太多 CPU|关键|在有限的时间内，Office 外接程序使用了超过 90% 的 CPU 资源。|
|15 |由于字符串搜索超时，外接程序已被禁用||Outlook 外接程序搜索电子邮件的主题行和消息，以确定是否应使用正则表达式来显示它们。“文件”列中列出的 Outlook 外接程序已被 Outlook 禁用，因为它在尝试匹配正则表达式时超时多次。|
|18 |外接程序已成功关闭||Office应用程序能够成功关闭Office加载项。|
|19|外接程序遇到运行时错误|关键|Office 外接程序遇到一个导致它失败的问题。 有关详细信息，请使用遇到错误的计算机上的 Windows 事件查看器查看“Microsoft Office 通知”日志。|
|20|外接程序未能验证许可|关键|无法验证 Office 外接程序的许可信息，且其可能已过期。 有关详细信息，请使用遇到错误的计算机上的 Windows 事件查看器查看“Microsoft Office 通知”日志。|

有关详细信息，请参阅[部署遥测仪表板](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15))和[使用遥测日志排查 Office 文件和自定义解决方案](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)。

## <a name="design-and-implementation-techniques"></a>设计和实现技术

尽管 CPU 和内存使用率的资源限制、故障容忍度以及 UI 无响应仅适用于在富客户端上运行的 Office 外接程序，但如果您希望外接程序在所有支持性客户端和设备上的性能都令人满意，优化这些资源和电池的使用情况仍然是头等大事。 如果您的外接程序要执行长时间运行的操作或处理大型数据集，则优化尤为重要。 以下列表建议了一些技术，可以将 CPU 密集型或数据密集型操作分解为较小的区块，这样加载项可以避免过多的资源消耗，并且 Office 应用程序可以保持响应。

- 在外接程序需要从无限制的数据集中读取大量数据的情况下，您可以在从表格中读取数据时应用分页，或者减小每次短暂读取操作中的数据大小，而不是试图在一次操作中完成全部读取。 可以通过全局对象的 [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) 方法来限制输入和输出的持续时间。 它还会处理定义区块中的数据，来代替随机无限数据。 另一个选项是 [使用异步来处理](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) 你的承诺。

- 如果外接程序使用大量占用 CPU 的算法来处理大量数据，则您可以使用 Web Worker 在后台执行长时间运行的任务，同时在前台运行单独的脚本，例如在用户界面中显示进度。Web Worker 不会阻止用户活动并允许 HTML 页面保持响应能力。有关 Web Worker 的示例，请参阅 [Web Worker 的基本信息](https://www.html5rocks.com/tutorials/workers/basics/)。有关 Web Worker API 的详细信息，请参阅 [Web Worker](https://developer.mozilla.org/docs/Web/API/Web_Workers_API)。

- 如果外接程序使用大量占用 CPU 的算法，但您可以将数据输入或输出划分成较小的集合，则可以考虑创建一个 Web 服务，将数据传递给该 Web 服务以减轻 CPU 负担，然后等待异步回调。

- 针对预期的最大数据量测试加载项，并限制加载项处理的数据量不得超过此限制。

### <a name="performance-improvements-with-the-application-specific-apis"></a>使用特定于应用程序的 API 的性能改进

Using the [application-specific API model](../develop/application-specific-api-model.md)中的性能提示提供有关将应用程序特定的 API 用于 Excel、OneNote、Visio 和 Word 的指南。 总之，你应该：

- [仅加载必要的属性](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended)。
- [最大程度地减少同步呼叫 () 数](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls)。 阅读 [避免在循环中使用 context.sync 方法](correlated-objects-pattern.md) ，详细了解如何在代码中管理 `sync` 调用。
- [尽量减少创建的代理对象数](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created)。 还可以取消跟踪代理对象，如下一节中所述。

#### <a name="untrack-unneeded-proxy-objects"></a>取消跟踪不需要的代理对象

[代理对象](../develop/application-specific-api-model.md#proxy-objects) 在调用之前一 `RequestContext.sync()` 直位于内存中。 大型批处理操作可能会生成许多代理对象，加载项只需用到这些对象一次，并且可以在批处理执行之前从内存中释放。

`untrack()`方法从内存中释放对象。 此方法在许多特定于应用程序的 API 代理对象上实现。 使用 对象完成加载项后调用，在使用大量代理对象时，应该会获得 `untrack()` 明显的性能优势。

> [!NOTE]
> `Range.untrack()` 是 [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove_object_) 的快捷方式。 任何代理对象都可以通过从上下文中的跟踪对象列表中删除它来取消跟踪。

下面的Excel示例使用数据填充选定区域，一次填充一个单元格。 将值添加到单元格后，表示该单元格的区域将被取消跟踪。 在选定的 10,000 到 20,000 个单元格区域运行此代码，首先使用 `cell.untrack()` 行，然后取消使用。 应会注意到，使用 `cell.untrack()` 行的代码比不使用的代码运行速度要快。 此外，可能还会注意到之后的响应时间更快，因为清理步骤花费的时间更少。

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // Call untrack() to release the range from memory.
            cell.untrack();
        }
    }

    await context.sync();
});
```

请注意，只有处理成千上万个对象时，才需要取消跟踪对象。 大多数外接程序不需要管理代理对象跟踪。

## <a name="see-also"></a>另请参阅

- [Office 加载项的隐私和安全](../concepts/privacy-and-security.md)
- [Outlook 外接程序的激活限制和 JavaScript API](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [使用 Excel JavaScript API 优化性能](../excel/performance.md)
