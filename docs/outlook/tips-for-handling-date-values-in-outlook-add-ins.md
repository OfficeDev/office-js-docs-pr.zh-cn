---
title: 处理 Outlook 加载项中的日期值
description: Office JavaScript API 对大多数存储和检索日期和时间使用 JavaScript Date 对象。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 49de8db712400e006dc919e9ad62ae6cbaaa11cf
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713075"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>处理 Outlook 加载项中的日期值的提示

Office JavaScript API 对大多数存储和检索日期和时间使用 JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) 对象。

该 `Date` 对象提供 [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) 和 [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp) 等方法，这些方法根据通用协调时间 (UTC) 时间返回请求的日期或时间值。

该 `Date` 对象还提供其他方法，例如 [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) 和 [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp)，它们根据“本地时间”返回请求的日期或时间。

“本地时间”概念很大程度上取决于客户端计算机上的浏览器和操作系统。 例如，在基于 Windows 的客户端计算机上运行的大多数浏览器上，JavaScript 调用 `getDate`会根据客户端计算机上 Windows 中设置的时区返回日期。

以下示例在本地时间创建一个 `Date` 对象 `myLocalDate` ，并调用 `toUTCString` 将该日期转换为 UTC 中的日期字符串。

```js
// Create and get the current date represented 
// in the client computer time zone.
const myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

虽然可以使用 JavaScript `Date` 对象来获取基于 UTC 或客户端计算机时区的日期或时间值，但 **Date** 对象在一个方面是有限的 - 它不提供返回任何其他特定时区的日期或时间值的方法。 例如，如果客户端计算机设置为在东部标准时间 (EST) 上，则没有 `Date` 方法允许你获取除 EST 或 UTC 以外的小时值，例如太平洋标准时间 (PST) 。

## <a name="date-related-features-for-outlook-add-ins"></a>Outlook 加载项的日期相关功能

如果使用 Office JavaScript API 处理在 Outlook 富客户端中运行的 Outlook 加载项中以及在 Outlook 网页版 或移动设备中运行的日期或时间值，上述 JavaScript 限制对你有含义。

### <a name="time-zones-for-outlook-clients"></a>Outlook 客户端的时区

为清楚起见，让我们先定义要讨论的时区。

|**时区**|**说明**|
|:-----|:-----|
|客户端计算机时区|这在客户端计算机的操作系统上设置。 大多数浏览器使用客户端计算机时区来显示 JavaScript `Date` 对象的日期或时间值。<br/><br/>Outlook 富客户端使用此时区在用户界面中显示日期或时间值。 <br/><br/>例如，在运行 Windows 的客户端计算机上，Outlook 将使用 Windows 上设置的时区作为本地时区。 在 Mac 上，如果用户更改客户端计算机上的时区，Outlook on Mac 也会提示用户更新 Outlook 中的时区。|
|Exchange 管理中心 (EAC) 时区|用户在首次登录Outlook 网页版或移动设备时 (和首选语言) 设置此时区值。 <br/><br/>Outlook 网页版和移动设备使用此时区在用户界面中显示日期或时间值。|

由于 Outlook 富客户端使用客户端计算机时区，并且Outlook 网页版和移动设备的用户界面使用 EAC 时区，因此，在 Outlook 富客户端和Outlook 网页版或移动设备中运行时，为同一邮箱安装的同一加载项的本地时间可能会有所不同。 作为 Outlook 外接程序开发人员，您应该正确输入和输出日期值，以便那些值始终与用户期望的相应客户端上的时区保持一致。

### <a name="date-related-api"></a>日期相关的 API

以下是 Office JavaScript API 中支持日期相关功能的属性和方法。

|API 成员|时区表示形式|Outlook 富客户端的示例|Outlook 网页版或移动设备中的示例|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#outlook-office-userprofile-timezone-member)|在 Outlook 富客户端中，此属性返回客户端计算机时区。 在Outlook 网页版和移动设备中，此属性返回 EAC 时区。 |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 和 [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|其中每个属性都返回一个 JavaScript `Date` 对象。 此`Date`值是正确的 UTC，如以下示例所示 - `myUTCDate` 在 Outlook 富客户端、Outlook 网页版和移动设备中具有相同的值。<br/><br/>`const myDate = Office.mailbox.item.dateTimeCreated;`<br/>`const myUTCDate = myDate.getUTCDate;`<br/><br/>但是，调用`myDate.getDate`在客户端计算机时区中返回一个日期值，该值与用于在 Outlook 富客户端接口中显示日期时间值的时区一致，但可能不同于Outlook 网页版和移动设备在其用户界面中使用的 EAC 时区。|如果项是在上午 9 点 UTC 创建的：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果项在 UTC 上午 11 点修改：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` 返回 6am EST。|如果项创建时间为上午 9 点 UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果项在 UTC 上午 11 点修改：<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` 返回 6am EST。<br/><br/>请注意，如果您想要在用户界面中显示创建或修改时间，要首先将时间转换为 PST 以与用户界面的其余部分保持一致。|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)|每个 _开始_ 和 _结束_ 参数都需要一个 JavaScript `Date` 对象。 无论 Outlook 富客户端或Outlook 网页版或移动设备的用户界面中使用的时区如何，参数都应为 UTC 正确。|如果约会表单的开始和结束时间是上午 9 点 UTC 和上午 11 点 UTC，则应确保 `start` 和 `end` 参数为 UTC 正确，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>|如果约会表单的开始和结束时间是上午 9 点 UTC 和上午 11 点 UTC，则应确保 `start` 和 `end` 参数为 UTC 正确，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>日期相关应用场景的帮助程序方法

如前面部分所述，由于 Outlook 富客户端上用户在Outlook 网页版或移动设备中的“本地时间”可能有所不同，但 JavaScript **Date** 对象仅支持转换为客户端计算机时区或 UTC，因此 Office JavaScript API 提供了两种帮助程序方法：[Office.context.mailbox.convertToLocalClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 和 [Office.context.mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)。

在 Outlook 富客户端、Outlook 网页版和移动设备中，这些帮助程序方法负责处理以下两种日期相关方案所需的日期或时间，从而增强外接程序的不同客户端的“写入一次”。

### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>应用场景 A：显示项创建时间或修改时间

如果在用户界面中显示项创建时间 (`Item.dateTimeCreated`) 或修改时间 (`Item.dateTimeModified`，请首先使用 `convertToLocalClientTime` 这些属性提供的对象来 `Date` 在适当的本地时间获取字典表示形式。 然后显示字典日期的各个部分。 下面是此方案的示例。

```js
// This date is UTC-correct.
const myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
const myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

请注意，`convertToLocalClientTime`请注意 Outlook 富客户端与Outlook 网页版或移动设备之间的差异：

- 如果 `convertToLocalClientTime` 检测到当前应用程序是富客户端，则该方法将表示形式转换为 `Date` 同一客户端计算机时区的字典表示形式，这与富客户端用户界面的其余部分一致。

- 如果`convertToLocalClientTime`检测到当前应用程序Outlook 网页版或移动设备，则该方法会将 UTC 正确的`Date`表示形式转换为 EAC 时区中的字典格式，这与其他Outlook 网页版或移动设备用户界面一致。

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>应用场景 B：在新的约会窗体中显示开始日期和结束日期

如果以输入形式获取在本地时间表示的日期时间值的不同部分，并且想要在约会窗体中将此字典输入值作为开始或结束时间提供，请首先使用 `convertToUtcClientTime` 帮助程序方法将字典值转换为 UTC 正确的 `Date` 对象。

在以下示例中，假定 `myLocalDictionaryStartDate` 和 `myLocalDictionaryEndDate` 是从用户获得的采用字典格式的日期时间值。 这些值基于本地时间，依赖于客户端平台。

```js
const myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
const myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

结果值 `myUTCCorrectStartDate` 和 `myUTCCorrectEndDate` 采用 UTC 格式。 然后将这些 `Date` 对象作为方法的 _“开始_ ”和 _“结束_ ”参数的 `Mailbox.displayNewAppointmentForm` 参数传递，以显示新的约会表单。

请注意，`convertToUtcClientTime`请注意 Outlook 富客户端与Outlook 网页版或移动设备之间的差异：

- 如果 `convertToUtcClientTime` 检测到当前应用程序是 Outlook 富客户端，则该方法只需将字典表示形式转换为 `Date` 对象。 此 `Date` 对象是 UTC 正确的对象，如预期的那样 `displayNewAppointmentForm`。

- 如果`convertToUtcClientTime`检测到当前应用程序Outlook 网页版或移动设备，则该方法会将 EAC 时区中表示的日期和时间值的字典格式转换为`Date`对象。 此 `Date` 对象是 UTC 正确的对象，如预期的那样 `displayNewAppointmentForm`。

## <a name="see-also"></a>另请参阅

- [部署和安装 Outlook 加载项以进行测试](testing-and-tips.md)
