---
title: 处理 Outlook 加载项中的日期值
description: Office JavaScript API 将 JavaScript Date 对象用于大多数日期和时间的存储和检索。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 48cbc407e21e377ed64dc873574d938b136bfd22
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292564"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>处理 Outlook 加载项中的日期值的提示

Office JavaScript API 将 JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) 对象用于大多数日期和时间的存储和检索。 

该 `Date` 对象提供诸如 [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [GetUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)和 [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp)的方法，这将根据 UTC) 时间 (的通用协调时间返回请求的日期或时间值。

该 `Date` 对象还提供了其他方法，如 [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、 [GetHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、 [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)和 [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp)，这将根据 "本地时间" 返回请求的日期或时间。

“本地时间”概念很大程度上取决于客户端计算机上的浏览器和操作系统。 例如，在基于 Windows 的客户端计算机上运行的大多数浏览器上，基于在 `getDate` 客户端计算机上的 windows 中设置的时区，对进行的 JavaScript 调用将返回一个日期。

下面的示例将创建 `Date` 一个 `myLocalDate` 本地时间对象，并调用将 `toUTCString` 该日期转换为 UTC 格式的日期字符串。

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

虽然您可以使用 JavaScript `Date` 对象来获取基于 UTC 或客户端计算机时区的日期或时间值，但在一个方面限制 **日期** 对象，但它不会提供返回任何其他特定时区的日期或时间值的方法。 例如，如果您的客户端计算机设置为在东部标准时间 (EST) 中，则没有任何 `Date` 方法允许您获取除 EST 或 UTC 之外的小时值，例如太平洋标准时间 (PST) 。


## <a name="date-related-features-for-outlook-add-ins"></a>Outlook 加载项的日期相关功能

当您使用 Office JavaScript API 在 outlook 富客户端中运行的 Outlook 外接程序中，以及在 Outlook 网页版或移动设备上运行时，上述 JavaScript 限制会给您暗示。


### <a name="time-zones-for-outlook-clients"></a>Outlook 客户端的时区

为清楚起见，让我们先定义要讨论的时区。

|**时区**|**说明**|
|:-----|:-----|
|客户端计算机时区|这在客户端计算机的操作系统上设置。 大多数浏览器都使用客户端计算机时区显示 JavaScript 对象的日期或时间值 `Date` 。<br/><br/>Outlook 富客户端使用此时区在用户界面中显示日期或时间值。 <br/><br/>例如，在运行 Windows 的客户端计算机上，Outlook 将使用 Windows 上设置的时区作为本地时区。 在 Mac 上，如果用户更改了客户端计算机上的时区，Mac 上的 Outlook 也会提示用户在 Outlook 中更新时区。|
|Exchange 管理中心 (EAC) 时区|用户将此时区值设置 (，首次登录到 web 或移动设备上的 Outlook 时) 首选语言。 <br/><br/>Web 和移动设备上的 Outlook 使用此时区在用户界面中显示日期或时间值。|

由于 Outlook 富客户端使用的是客户端计算机时区，web 上的 Outlook 和移动设备上的用户界面使用 EAC 时区，因此在 Outlook 富客户端和 web 或移动设备上的 Outlook 中运行相同的加载项时，为同一邮箱安装的本地时间可能会有所不同。 作为 Outlook 外接程序开发人员，您应该正确输入和输出日期值，以便那些值始终与用户期望的相应客户端上的时区保持一致。


### <a name="date-related-api"></a>日期相关的 API

以下是 Office JavaScript API 中支持日期相关功能的属性和方法。

**API 成员**|**时区表示形式**|**Outlook 富客户端中的示例**|**Web 或移动设备上的 Outlook 中的示例**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|在 Outlook 富客户端中，此属性返回客户端计算机时区。 在 web 和移动设备上的 Outlook 中，此属性返回 EAC 时区。 |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|这些属性中的每一个都返回一个 JavaScript `Date` 对象。 此 `Date` 值是 UTC 正确的，如以下示例中所示- `myUTCDate` outlook 富客户端中的值相同，outlook 网页版和移动设备上具有相同的值。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>但是，调用将  `myDate.getDate` 返回客户端计算机时区中的 date 值，这与用于在 outlook 富客户端界面中显示日期时间值的时区一致，但可能不同于 web 和移动设备上的用户界面中使用的 EAC 时区。|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` 返回 6am EST。|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` 返回 6am EST。<br/><br/>请注意，如果您想要在用户界面中显示创建或修改时间，要首先将时间转换为 PST 以与用户界面的其余部分保持一致。
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|每个  _Start_ 和 _End_ 参数都需要一个 JavaScript `Date` 对象。 参数应采用 UTC 格式，无论 Outlook 富客户端的用户界面中使用的时区或 web 或移动设备上的 Outlook 中使用的是何种时区。|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>日期相关应用场景的帮助程序方法


如前面的部分中所述，由于 web 或移动设备上的 Outlook 中的用户的 "本地时间" 在 Outlook 富客户端中可能不同，但 JavaScript **Date** 对象支持仅转换为客户端计算机时区或 UTC，而 OFFICE JavaScript API 提供了两种帮助程序方法： [convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 和 [convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)。

这些帮助程序方法负责在 Outlook 富客户端、web 和移动设备上的 outlook 富客户端、Outlook 网页版和移动设备上为以下两个与日期相关的方案处理日期或时间的任何需求，从而为外接程序的不同客户端加强 "一次性写入"。


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>应用场景 A：显示项创建时间或修改时间

如果要在用户界面中显示项目创建时间 (`Item.dateTimeCreated`) 或修改时间 (`Item.dateTimeModified` ，请首先使用 `convertToLocalClientTime` 来转换 `Date` 这些属性提供的对象，以在适当的当地时间获取字典表示形式。 然后显示字典日期的各个部分。 下面是此应用场景的一个示例：


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

请注意， `convertToLocalClientTime` 解决 outlook 富客户端和 web 或移动设备上的 outlook 之间的差异：


- 如果 `convertToLocalClientTime` 检测到当前应用程序是富客户端，则此方法会将 `Date` 表示形式转换为同一客户端计算机时区中的字典表示形式，与富客户端用户界面的其余部分一致。
    
- 如果 `convertToLocalClientTime` 检测到当前应用程序是在 web 或移动设备上的 outlook，则此方法会将 UTC `Date` 格式的表示形式转换为 EAC 时区中的词典格式，与 Outlook 在 web 或移动设备用户界面上的其余部分保持一致。
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>应用场景 B：在新的约会窗体中显示开始日期和结束日期

如果要获取以本地时间表示的日期时间值的输入不同部分，并且希望在约会窗体中提供此字典输入值作为开始时间或结束时间，请首先使用 `convertToUtcClientTime` helper 方法将字典值转换为与 UTC 正确关联的 `Date` 对象。

在以下示例中，假定 `myLocalDictionaryStartDate` 和 `myLocalDictionaryEndDate` 是从用户获得的采用字典格式的日期时间值。 这些值基于本地时间，具体取决于客户端平台。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

结果值 `myUTCCorrectStartDate` 和 `myUTCCorrectEndDate` 采用 UTC 格式。 然后，将这些 `Date` 对象作为参数传递给方法的 _Start_ 和 _End_ 参数 `Mailbox.displayNewAppointmentForm` ，以显示新的约会窗体。

请注意， `convertToUtcClientTime` 解决 outlook 富客户端和 web 或移动设备上的 outlook 之间的差异：


- 如果 `convertToUtcClientTime` 检测到当前应用程序是 Outlook 富客户端，则此方法只是将字典表示形式转换为 `Date` 对象。 此 `Date` 对象按所需的 UTC 格式正确 `displayNewAppointmentForm` 。
    
- 如果 `convertToUtcClientTime` 检测到当前应用程序是在 web 或移动设备上的 Outlook，则此方法会将在 EAC 时区中表示的日期和时间值的词典格式转换为 `Date` 对象。 此 `Date` 对象按所需的 UTC 格式正确 `displayNewAppointmentForm` 。
    
## <a name="see-also"></a>另请参阅

- [部署和安装 Outlook 加载项以进行测试](testing-and-tips.md)
