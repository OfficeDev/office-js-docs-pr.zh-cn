---
title: Office 常用 API 错误代码
description: 本文记录了在使用 Office 通用 API 时可能会遇到的错误消息。
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d77b4c0c458e11da0057f06a5088ef8a28e4ccd2
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092978"
---
# <a name="office-common-api-error-codes"></a>Office 常用 API 错误代码

本文记录了在使用 Common API 模型时可能会遇到的错误消息。 这些错误代码不适用于特定于应用程序的 API，例如 Excel JavaScript API 或 Word JavaScript API。

请参阅 [API 模型](../develop/understanding-the-javascript-api-for-office.md#api-models) ，详细了解通用 API 与特定于应用程序的 API 模型之间的差异。

## <a name="error-codes"></a>错误代码

下表列出了显示的错误代码、名称和消息，及其指示的条件。

|Error.code|Error.name|Error.message|条件|
|:-----|:-----|:-----|:-----|
|1000|强制类型无效|不支持指定的强制类型|Office 应用程序不支持强制类型。  (例如，Excel.) 不支持 OOXML 和 HTML 强制类型|
|1001|数据读取错误|不支持当前所选内容。|不支持用户当前所选内容（即与受支持的强制类型有所不同）。|
|1002|强制类型无效|指定的强制类型与此绑定类型不兼容。|解决方案开发人员提供了强制类型和绑定类型的不兼容组合。|
|1003|数据读取错误|指定的 rowCount 或 columnCount 值无效。|用户提供无效的列或行计数。|
|1004|数据读取错误|当前选定内容与指定的强制类型不兼容。|当前选定内容不支持此应用程序指定的强制类型。|
|1005|数据读取错误|指定的 startRow 或 startColumn 值无效。|用户提供的 startRow 或 startCol 值无效。|
|1006|数据读取错误|如果表格包含合并单元格，则坐标参数不能用于强制类型“Table”。|用户尝试获取非统一表格（即包含合并单元格的表格）的部分数据。 |
|1007|数据读取错误|文档太大。|用户尝试获取大于当前支持大小的文档。|
|1008|数据读取错误|请求的数据集太大。|用户请求读取超出 Office 应用程序定义的数据限制的数据。|
|1009|数据读取错误|不支持指定的文件类型。|用户发送的文件类型无效。|
|2000|数据写入错误|不支持提供的数据对象类型。 |提供了不受支持的数据对象。|
|2001|数据写入错误|无法写入当前所选内容。|The user's current selection is not supported for a write operation. (For example, when the user selects an image.)|
|2002|数据写入错误|提供的数据对象与当前所选内容的形状或尺寸不兼容。|选择了多个单元格（且所选内容的形状与数据的形状不匹配）。 选择了多个单元格（且所选内容的尺寸与数据的尺寸不匹配）。|
|2003|数据写入错误|设置操作失败，因为提供的数据对象将覆盖数据。|选择了单个单元格，且提供的数据对象将覆盖工作簿中的数据。|
|2004|数据写入错误|提供的数据对象与当前所选内容的大小不匹配。|用户提供的对象大于当前所选内容的大小。|
|2005|数据写入错误|指定的 startRow 或 startColumn 值无效。|用户提供的 startRow 或 startCol 值无效。|
|2006|无效格式错误|指定数据对象的格式无效。|解决方案开发人员提供了无效的 HTML 或 OOXML 字符串、格式错误的 HTML 字符串或无效的 OOXML 字符串。|
|2007|数据对象无效|指定数据对象的类型与当前所选内容不兼容。|解决方案开发人员提供的数据对象与指定的强制类型不兼容。|
|2008|数据写入错误|TBD|TBD|
|2009|数据写入错误|指定的数据对象太大。|用户尝试将数据设置为超出 Office 应用程序定义的数据限制。|
|2010|数据写入错误|如果表格包含合并单元格，则坐标参数不能用于强制类型"Table"。|用户尝试设置非统一表格（即包含合并单元格的表格）的部分数据。|
|3000|绑定创建错误|无法绑定到当前所选内容。|The user's selection is not supported for binding. (For example, the user is selecting an image or other non-supported object.)|
|3001|绑定创建错误|TBD|TBD|
|3002|无效绑定错误|指定绑定不存在。|开发人员尝试绑定到不存在或已删除的绑定。|
|3003|绑定创建错误|不支持非连续的所选内容。|用户进行多重选择。|
|3004|绑定创建错误|无法使用当前所选内容和指定绑定类型创建绑定。|There are several conditions under which this might happen. Please see the "Binding creation error conditions" section later in this article.|
|3005|绑定操作无效|此绑定类型不支持操作。|开发人员在不是强制 `table`类型的绑定类型上发送添加行或添加列操作。|
|3006|绑定创建错误|命名项不存在。|The named item cannot be found. No content control or table with that name exists.|
|3007|绑定创建错误|找到具有相同名称的多个对象。|碰撞错误：存在多个名称相同的内容控件，并在冲突时失败设置为 `true`。|
|3008|绑定创建错误|指定的绑定类型与提供的命名项不兼容。|命名项不能绑定为类型。 例如，内容控件包含文本，但开发人员尝试使用强制类型 `table`进行绑定。|
|3009|绑定操作无效|不支持绑定类型。|用于向后兼容性。|
|3010|不受支持的绑定操作|所选内容需为表格格式。 将数据格式化为表格并再次尝试。|开发人员尝试对强制类型的`matrix`数据使用`addRowsAsync`对象或`deleteAllDataValuesAsync`方法`TableBinding`。|
|4000|读取设置错误|指定设置名称不存在。|提供了不存在的设置名称。|
|4001|保存设置错误|无法保存设置。|无法保存设置。|
|4002|设置过期错误|无法保存设置，因为设置已过期。|设置已过期，开发人员指示不要覆盖设置。|
|5000|设置过期错误|不支持此操作。|当前 Office 应用程序不支持此操作。 例如， `document.getSelectionAsync` 从 Outlook 调用。|
|5001|内部错误|发生内部错误。|请参阅内部错误情况，可能有以下几个原因。<br/><table><tr><td>另一个共享工作簿的用户正在使用的加载项几乎在同一时间创建了一个绑定，您的加载项需要重新尝试绑定。</tr></td><tr><td>出现未知错误。</tr></td><tr><td>操作失败。</tr></td><tr><td>访问被拒绝，因为用户不是已授权角色的成员。</tr></td><tr><td>访问被拒绝，因为要求安全、加密的通信。</tr></td><tr><td>数据已过时，用户需要确认启用查询以刷数据。</tr></td><tr><td>已超出网站集 CPU 配额。</tr></td><tr><td>已超出网站集的内存配额。</tr></td><tr><td>已超出会话内存配额。</tr></td><tr><td>工作簿处于无效状态，无法执行该操作。</tr></td><tr><td>会话因不活动而超时，用户需要重新加载工作簿。</tr></td><tr><td>已超出每个用户允许的会话数最大值。</tr></td><tr><td>操作被用户取消。</tr></td><tr><td>因为时间太长，无法完成该操作。</tr></td><tr><td>该请求无法完成，需要重试。</tr></td><tr><td>产品的试用期已过。</tr></td><tr><td>会话因不活动而超时。</tr></td><tr><td>用户不具有在指定范围内执行该操作的权限。</tr></td><tr><td>用户的区域设置与当前协作会话不匹配。</tr></td><tr><td>用户已断开连接，必须刷新或重新打开工作簿。</tr></td><tr><td>工作表中不存在请求的范围。</tr></td><tr><td>用户没有编辑该工作簿的权限。</tr></td><tr><td>工作簿已锁定，无法编辑。</tr></td><tr><td>会话无法自动保存工作簿。</tr></td><tr><td>会话无法刷新其在工作簿文件上的锁定。</tr></td><tr><td>无法处理请求，需要重试。</tr></td><tr><td>无法验证用户的登录信息，必须重新输入。</tr></td><tr><td>用户访问被拒绝。</tr></td><tr><td>需要更新共享工作簿。</tr></td></table>|
|5002|权限被拒绝|当前文档模式不允许请求的操作。|解决方案开发人员提交一组操作，但文档模式不允许进行修改，如"限制编辑"。|
|5003|事件注册错误|当前对象不支持指定事件类型。|解决方案开发人员尝试对不存在的事件注册或取消注册处理程序。|
|5004|无效的 API 调用|当前上下文中无效的 API 调用。|对上下文进行了无效调用，例如，尝试在 Excel 中使用 `CustomXMLPart` 对象。|
|5005|数据过期|操作失败，因为服务器上的数据已过期。|需要刷新服务器上的数据。|
|5006|会话超时|文档会话超时。重新加载文档。 |会话已超时。|
|5007|无效的 API 调用|当前上下文中不支持枚举。|当前上下文中不支持枚举。|
|5009|权限被拒绝|访问被拒绝|加载项没有调用特定 API 的权限。|
|5012|会话无效或超时|Office 浏览器会话已过期或无效。 若要继续操作，请刷新页面。|Office 客户端和服务器之间的会话已过期，或你的计算机上的日期、时间或时区不正确。|
|6000|无效节点|未找到指定节点。|`CustomXmlPart`找不到节点。|
|6100|自定义 XML 错误|自定义 XML 错误|无效的 API 调用。|
|7000|ID 无效|指定 ID 不存在。|ID 无效。|
|7001|导航无效|导航不支持对象位置。|The user can find the object, but cannot navigate to it. (For example, in Word, the binding is to the header, footer, or a comment.)|
|7002|导航无效|对象已锁定或受保护。|用户尝试导航到被阻止或受保护的范围。|
|7004|导航无效|操作失败，索引已超出范围。|用户尝试导航到索引，但超出范围。|
|8000|参数缺失|We couldn't format the table cell because some parameter values are missing. Double-check the parameters and try again.|The cellFormat method is missing some parameters. For example, there are missing cells, format, or tableOptions parameters.|
|8010|值无效|One or more of the cells parameters have values that aren't allowed. Double-check the values and try again.|The common cells reference enumeration is not defined. For example, All, Data, Headers.|
|8011|值无效|One or more of the tableOptions parameters have values that aren't allowed. Double-check the values and try again.|tableOptions 中的某个值无效。|
|8012|值无效|One or more of the format parameters have values that aren't allowed. Double-check the values and try again.|格式中的某个值无效。|
|8020|超出范围|The row index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of rows.|行索引大于表格的最大行索引或小于 0。|
|8021|超出范围|The column index value is out of the allowed range. Use a positive value (0 or higher) that's less than the number of columns.|列索引大于表格最大列索引或小于 0。|
|8022|超出范围|值超出允许的范围。|格式中的某些值超出支持的范围。|
|9016|权限被拒绝|权限被拒绝|访问被拒绝。|
|9020|泛型响应错误|发生内部错误。|引用内部错误条件，由于任意数量的原因，可能会发生此情况。|
|9021|保存错误|尝试在服务器上保存项时出现连接错误。|无法保存该项。 这可能是由于在 Outlook 桌面版中使用联机模式时出现服务器连接错误，或者是由于尝试重新保存从 Exchange 服务器中删除的草稿项。|
|9022|不同存储区中的消息错误|无法检索 EWS ID，因为消息保存在另一个存储区中。|无法检索当前邮件的 EWS ID，因为邮件可能已移动，或者发送邮箱可能已更改。|
|9041|网络错误|用户不再连接到网络。 请检查网络连接并重试。|用户不再具有网络或 Internet 访问权限。|
|9043|不支持的附件类型|不支持附件类型。|API 不支持附件类型。 例如， `item.getAttachmentContentAsync` 如果附件是带丰富文本格式的嵌入式图像，或者它是电子邮件或日历项以外的项目类型，则引发此错误 (如联系人或任务项) 。|
|12002|*不适用。*|*不适用。*|下列一种含义：<br> - 传递给 `displayDialogAsync` 的 URL 没有对应的页面。<br> - 传递给 `displayDialogAsync` 的页面已加载，但对话框定向到找不到或无法加载的页面，或者已定向到使用无效语法的 URL。 在对话框中引发并在主机页面中触发 `DialogEventReceived` 事件。|
|12003|*不适用。*|*不适用。*|对话框定向到使用 HTTP 协议的 URL。 必须使用 HTTPS。 在对话框中引发并在主机页面中触发 `DialogEventReceived` 事件。|
|12004|*不适用。*|*不适用。*|传递给 `displayDialogAsync` 的 URL 的域不受信任。 此域必须与主机页的域相同（包括协议和端口号）。 由 `displayDialogAsync` 的调用引发。|
|12005|*不适用。*|*不适用。*|传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。 必须使用 HTTPS。 由 `displayDialogAsync` 的调用引发。 （在 Office 的某些版本中，返回 12005 的错误消息与返回 12004 错误消息是相同的。）|
|12006|*不适用。*|*不适用。*|对话框已关闭，通常是因为用户选择了 **X** 按钮。 在对话框中引发并在主机页面中触发 `DialogEventReceived` 事件。|
|12007|*不适用。*|*不适用。*|已从此主机窗口打开了一个对话框。 主机窗口（如任务窗格）一次只能打开一个对话框。 由 `displayDialogAsync` 的调用引发。|
|12009|*不适用。*|*不适用。*|用户已选择忽略对话框。 联机版本的 Office 中可能会发生此错误，用户可能会选择不允许加载项显示对话框。 由 `displayDialogAsync` 的调用引发。|
|12011|*不适用。*|*不适用。*|用户的浏览器的配置方式是阻止弹出窗口。 如果浏览器是 Safari 且已配置为阻止弹出窗口，或者浏览器为 Edge Legacy，并且外接程序域与对话框尝试打开的域位于不同的安全区域，则Office web 版可能会发生此错误。 由 `displayDialogAsync` 的调用引发。|
|13nnn|*不适用。*|*不适用。*|查看 [getAccessToken 中错误的原因和处理](../develop/troubleshoot-sso-in-office-add-ins.md#causes-and-handling-of-errors-from-getaccesstoken)。|

## <a name="binding-creation-error-conditions"></a>绑定创建错误条件

When a binding is created in the API, indicate the binding type that you want to use. The following tables lists the binding types and the resulting binding behaviors that are expected.

### <a name="behavior-in-excel"></a>Excel 中的行为

下表汇总了 Excel 中的绑定行为。

|指定绑定类型|实际选择|行为|
|:-----|:-----|:-----|
|矩阵|单元格范围（包括表格和单个单元格范围内）|在所选单元格上创建类型的 `matrix` 绑定。 不得修改文档。|
|Matrix|单元格中选定的文本|在整个单元格上创建类型的 `matrix` 绑定。 不得修改文档。|
|Matrix|多重选择/选择无效（例如，用户选择了图片、对象或艺术字。）|无法创建绑定。|
|Table|单元格范围（包括单个单元格）|无法创建绑定。|
|Table|表格内单元格的范围（包括表格中单个单元格、整张表格或表格中单元格内的文本）|已在整张表格中创建绑定。|
|Table|表格中和表格外的半选定|无法创建绑定。|
|Table|单元格（而非表格）中选定的文本。|无法创建绑定。|
|Table|多重选择/选择无效（例如，用户选择了图片、对象、艺术字等。）|无法创建绑定。|
|文本|单元格范围|无法创建绑定。|
|文本|表格内的单元格范围|无法创建绑定。|
|文本|单个单元格|将创建类型的 `text` 绑定。|
|Text|表格内的单个单元格|将创建类型的 `text` 绑定。|
|Text|单元格中选定的文本|将创建整个单元格中类型的 `text` 绑定。|

### <a name="behavior-in-word"></a>Word 中的行为

下表汇总了 Word 中的绑定行为。

|指定绑定类型|实际选择|行为|
|:-----|:-----|:-----|
|矩阵|文本|无法创建绑定。|
|Matrix|整张表格|将创建类型的 `matrix` 绑定。文档已更改，内容控件必须包装表。 |
|矩阵|表格范围内|无法创建绑定。|
|Matrix|选择无效（例如，多个对象、无效对象等。）|无法创建绑定。|
|Table|文本|无法创建绑定。|
|Table|整张表格|将创建类型的 `text` 绑定。|
|表格|表格范围内|无法创建绑定。|
|Table|选择无效（例如，多个对象、无效对象等。）|无法创建绑定。|
|文本|整张表格|将创建类型的 `text` 绑定。|
|Text|表格范围内|无法创建绑定。|
|文本|多重选择|最后的选定内容必须与内容控件一同打包，并绑定到该控件。 将创建类型的 `text` 内容控件。|
|Text|选择无效（例如，多个对象、无效对象等。）|无法创建绑定。|

## <a name="see-also"></a>另请参阅

- [Office 加载项开发生命周期](../overview/office-add-ins.md)
- [了解 Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md)
- [使用特定于应用程序的 JavaScript API 进行错误处理](../testing/application-specific-api-error-handling.md)
- [排查单一登录 (SSO) 错误消息](../develop/troubleshoot-sso-in-office-add-ins.md)
- [排查 Office 加载项中的开发错误](../testing/troubleshoot-development-errors.md)
