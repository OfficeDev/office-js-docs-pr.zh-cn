---
title: 处理 Outlook 加载项中的日期值
description: JavaScript API Office JavaScript Date 对象对日期和时间进行大部分存储和检索。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 46be9e7e3c952d08addcf8ef761a259f8c0d1d84c1bc3b0bb61cbb40c07ce35b
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093316"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>处理 Outlook 加载项中的日期值的提示

JavaScript API Office JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp)对象用于大多数日期和时间的存储和检索。 

该对象提供 `Date` getUTCDate、getUTCHour、getUTCMinutes 和[toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp)[](https://www.w3schools.com/jsref/jsref_getutcdate.asp)[](https://www.w3schools.com/jsref/jsref_getutchours.asp)等方法，这些方法根据协调世界时 (UTC) 时间返回请求的日期或时间值[](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)。

对象 `Date` 还提供了其他方法，如 [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)和 [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp)，这些方法根据"本地时间"返回请求的日期或时间。

“本地时间”概念很大程度上取决于客户端计算机上的浏览器和操作系统。 例如，在基于 Windows 的客户端计算机上运行的多数浏览器上，JavaScript 调用 将基于客户端计算机上 Windows 中设置的 `getDate` 时区返回日期。

下面的示例在本地时间创建一个对象，并调用 `Date` `myLocalDate` `toUTCString` 以将日期转换为 UTC 格式的日期字符串。

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

虽然可以使用 JavaScript 对象获取基于 UTC 或客户端计算机时区的日期或时间值，但 Date 对象在一方面受到限制，它不提供用于返回任何其他特定时区的日期或时间值的方法。 `Date`  例如，如果客户端计算机设置为东部标准时间 (EST) ，则没有方法允许您获取除 EST 或 UTC 格式的小时值，例如太平洋标准时间 `Date` (PST) 。


## <a name="date-related-features-for-outlook-add-ins"></a>Outlook 加载项的日期相关功能

使用 Office JavaScript API 处理在 Outlook 富客户端、Outlook 网页版 或移动设备中运行的 Outlook 外接程序中的日期或时间值时，上述 JavaScript 限制有一定影响。


### <a name="time-zones-for-outlook-clients"></a>Outlook 客户端的时区

为清楚起见，让我们先定义要讨论的时区。

|**时区**|**说明**|
|:-----|:-----|
|客户端计算机时区|这在客户端计算机的操作系统上设置。 大多数浏览器使用客户端计算机时区来显示 JavaScript 对象的日期或时间 `Date` 值。<br/><br/>Outlook 富客户端使用此时区在用户界面中显示日期或时间值。 <br/><br/>例如，在运行 Windows 的客户端计算机上，Outlook 将使用 Windows 上设置的时区作为本地时区。 在 Mac 上，如果用户更改客户端计算机上的时区，Outlook Mac 上的用户也会提示用户更新 Outlook 时区。|
|Exchange 管理中心 (EAC) 时区|用户在首次登录 (移动设备时) 时区值Outlook 网页版首选语言。 <br/><br/>Outlook 网页版移动设备使用此时区在用户界面中显示日期或时间值。|

由于 Outlook 富客户端使用客户端计算机时区，并且 Outlook 网页版 和移动设备的用户界面使用 EAC 时区，因此，在 Outlook 富客户端和 Outlook 网页版 或移动设备中运行时，为同一邮箱安装的同一外接程序的本地时间可能不同。 作为 Outlook 外接程序开发人员，您应该正确输入和输出日期值，以便那些值始终与用户期望的相应客户端上的时区保持一致。


### <a name="date-related-api"></a>日期相关的 API

以下是支持日期相关功能的 Office JavaScript API 中的属性和方法。

|API 成员|时区表示形式|Outlook 富客户端的示例|移动设备Outlook 网页版中的示例|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#timeZone)|在 Outlook 富客户端中，此属性返回客户端计算机时区。 在Outlook 网页版移动设备中，此属性返回 EAC 时区。 |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|其中每个属性返回一个 JavaScript `Date` 对象。 此值为 UTC 格式，如以下示例所示 - 在富客户端、Outlook和移动设备Outlook 网页版 `Date` `myUTCDate` 值。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>但是，调用会返回客户端计算机时区中的日期值，该值与用于在 Outlook 富客户端界面中显示日期时间值的时区一致，但可能不同于 Outlook 网页版 和移动设备在其用户界面中使用的 EAC 时区。 `myDate.getDate`|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` 返回 6am EST。|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` 返回 6am EST。<br/><br/>请注意，如果您想要在用户界面中显示创建或修改时间，要首先将时间转换为 PST 以与用户界面的其余部分保持一致。|
|[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|每个  _Start 和_ _End_ 参数都需要一个 JavaScript `Date` 对象。 参数应采用 UTC 格式，而不考虑在富客户端、Outlook或移动设备Outlook 网页版使用的时区。|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>日期相关应用场景的帮助程序方法


如前面部分所述，由于 Outlook 网页版 或移动设备中的用户的"本地时间"在 Outlook 富客户端上可能不同，但 JavaScript **Date** 对象仅支持转换为客户端计算机时区或 UTC，因此 Office JavaScript API 提供了两个帮助程序方法 [：Office.context.mailbox.convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)和 [Office.context.mailbox.convertToUtcClientTime。](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)

这些帮助程序方法可处理以下两个与日期相关的方案（在 Outlook 富客户端、Outlook 网页版 和移动设备中）以不同方式处理日期或时间的任何需求，从而强化外接程序的不同客户端的"写入一次"。


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>应用场景 A：显示项创建时间或修改时间

如果要在用户界面中显示项目创建时间 () 或修改时间 (，则首先使用 转换这些属性提供的对象，以在适当的本地时间获取字典 `Item.dateTimeCreated` `Item.dateTimeModified` `convertToLocalClientTime` `Date` 表示形式。 然后显示字典日期的各个部分。 下面是此方案的一个示例。


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

请注意 `convertToLocalClientTime` ，需注意富客户端Outlook移动设备Outlook 网页版差异：


- 如果检测到当前应用程序是富客户端，则该方法将表示形式转换为同一客户端计算机时区中的字典表示形式，与富客户端用户界面的其余部分 `convertToLocalClientTime` `Date` 保持一致。
    
- 如果检测到当前应用程序是 Outlook 网页版 或移动设备，则该方法将采用 UTC 格式的表示形式转换为 EAC 时区中的字典格式，与 Outlook 网页版 或移动设备用户界面的其余部分一 `convertToLocalClientTime` `Date` 致。
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>应用场景 B：在新的约会窗体中显示开始日期和结束日期

如果要获取以本地时间表示的日期时间值的不同部分作为输入，并且希望将此字典输入值作为约会窗体中的开始时间或结束时间提供，请首先使用帮助程序方法将字典值转换为 UTC 正确的对象。 `convertToUtcClientTime` `Date`

在以下示例中，假定 `myLocalDictionaryStartDate` 和 `myLocalDictionaryEndDate` 是从用户获得的采用字典格式的日期时间值。 这些值基于本地时间，取决于客户端平台。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

结果值 `myUTCCorrectStartDate` 和 `myUTCCorrectEndDate` 采用 UTC 格式。 然后将这些 `Date` 对象作为 _方法的 Start_ 和 _End_ 参数的参数传递， `Mailbox.displayNewAppointmentForm` 以显示新的约会窗体。

请注意 `convertToUtcClientTime` ，需注意富客户端Outlook移动设备Outlook 网页版差异：


- 如果 `convertToUtcClientTime` 检测到当前应用程序是富Outlook，该方法只是将字典表示形式转换为 `Date` 对象。 此 `Date` 对象是 UTC 格式，如 预期的那样 `displayNewAppointmentForm` 。
    
- 如果检测到当前应用程序Outlook 网页版移动设备，该方法会将用 EAC 时区表示的日期和时间值的字典格式 `convertToUtcClientTime` 转换为 `Date` 对象。 此 `Date` 对象是 UTC 格式，如 预期的那样 `displayNewAppointmentForm` 。
    
## <a name="see-also"></a>另请参阅

- [部署和安装 Outlook 加载项以进行测试](testing-and-tips.md)
