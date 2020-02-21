---
title: 处理 Outlook 加载项中的日期值
description: 适用于 Office 的 JavaScript API 将 JavaScript Date 对象用于大多数日期和时间存储和检索。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5718839ebda433df6fb14886da34d734f81eb5f2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165884"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>处理 Outlook 加载项中的日期值的提示

适用于 Office 的 JavaScript API 将 JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) 对象用于大多数日期和时间存储和检索。 

该 **Date** 对象提供一些方法，如 [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、[getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) 和 [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp)，它根据协调世界时 (UTC) 时间返回请求的日期或时间值。

**Date** 对象还提供了其他方法，如 [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp)、[getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)、[getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) 和 [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp)，它根据本地时间返回请求的日期或时间。

“本地时间”概念很大程度上取决于客户端计算机上的浏览器和操作系统。例如，在大多数运行于基于 Windows 的客户端计算机的浏览器上，JavaScript 调用 **getDate** 根据客户端计算机上 Windows 中设置的时区返回日期。

下面的示例以本地时间创建 **Date** 对象 `myLocalDate`，然后调用 **toUTCString** 将该日期转换为 UTC 格式的日期字符串。

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

虽然可以使用 JavaScript **Date** 对象获取基于 UTC 或客户端计算机时区的日期或时间值，但是 **Date** 对象在一个方面受到限制：它不提供任何方法来返回任何其他特定时区的日期或时间值。 例如，如果客户端计算机设置为东部标准时间 (EST)，则无法使用任何 **Date** 方法来获取除 EST 和 UTC 外（如太平洋标准时间 (PST)）的小时值。


## <a name="date-related-features-for-outlook-add-ins"></a>Outlook 加载项的日期相关功能

当您使用适用于 Office 的 JavaScript API 在 outlook 富客户端中运行的 Outlook 外接程序中以及在 web 或移动设备上运行的 outlook 外接程序中的日期或时间值时，上述 JavaScript 限制将为您暗示。


### <a name="time-zones-for-outlook-clients"></a>Outlook 客户端的时区

为清楚起见，让我们先定义要讨论的时区。

|**时区**|**说明**|
|:-----|:-----|
|客户端计算机时区|这在客户端计算机的操作系统上设置。 大多数浏览器使用客户端计算机时区来显示 JavaScript **Date** 对象的日期或时间值。<br/><br/>Outlook 富客户端使用此时区在用户界面中显示日期或时间值。 <br/><br/>例如，在运行 Windows 的客户端计算机上，Outlook 将使用 Windows 上设置的时区作为本地时区。 在 Mac 上，如果用户更改了客户端计算机上的时区，Mac 上的 Outlook 也会提示用户在 Outlook 中更新时区。|
|Exchange 管理中心 (EAC) 时区|当用户首次登录到 web 或移动设备上的 Outlook 时，用户将设置此时区值（和首选语言）。 <br/><br/>Web 和移动设备上的 Outlook 使用此时区在用户界面中显示日期或时间值。|

由于 Outlook 富客户端使用的是客户端计算机时区，web 上的 Outlook 和移动设备上的用户界面使用 EAC 时区，因此在 Outlook 富 clie 中运行时，为同一邮箱安装的相同加载项的本地时间可能不同。nt 和在 Outlook 网页版或移动设备上。 作为 Outlook 外接程序开发人员，您应该正确输入和输出日期值，以便那些值始终与用户期望的相应客户端上的时区保持一致。


### <a name="date-related-api"></a>日期相关的 API

以下是适用于 Office 的 JavaScript API 中支持日期相关功能的属性和方法。

**API 成员**|**时区表示形式**|**Outlook 富客户端中的示例**|**Web 或移动设备上的 Outlook 中的示例**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|在 Outlook 富客户端中，此属性返回客户端计算机时区。 在 web 和移动设备上的 Outlook 中，此属性返回 EAC 时区。 |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 和 [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|上述每个属性返回 JavaScript **Date** 对象。 此**日期**值是正确的，如以下示例中所示- `myUTCDate` outlook 富客户端中具有相同的值，outlook 网页版和移动设备上具有相同的值。<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>但是，调用`myDate.getDate`将返回客户端计算机时区中的 date 值，这与用于在 outlook 富客户端界面中显示日期时间值的时区一致，但可能不同于 web 和移动设备上的用户界面中使用的 EAC 时区。|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` 返回 6am EST。|如果此项的创建时间是 9am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` 返回 4am EST。<br/><br/>如果此项的修改时间是 11am UTC：<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` 返回 6am EST。<br/><br/>请注意，如果您想要在用户界面中显示创建或修改时间，要首先将时间转换为 PST 以与用户界面的其余部分保持一致。
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|每个 _Start_ 和 _End_ 参数都需要一个 JavaScript **Date** 对象。 参数应采用 UTC 格式，无论 Outlook 富客户端的用户界面中使用的时区或 web 或移动设备上的 Outlook 中使用的是何种时区。|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>|如果约会窗体的开始和结束时间分别是 9am UTC 和 11am UTC，则应确保 `start` 和 `end` 参数都是 UTC 格式，这意味着：<br/><br/><ul><li>`start.getUTCHours` 返回 9am UTC</li><li>`end.getUTCHours` 返回 11am UTC</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>日期相关应用场景的帮助程序方法


如前面各节中所述，由于 web 或移动设备上的 Outlook 中的用户的 "本地时间" 在 Outlook 富客户端中可能不同，但 JavaScript **Date**对象支持仅转换为客户端计算机时区或 UTC，适用于 Office 的 javascript API 提供了两种帮助程序方法： [convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)和[convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)。

这些帮助程序方法负责在 Outlook 富客户端、web 和移动设备上的 outlook 富客户端、Outlook 网页版和移动设备上为以下两个与日期相关的方案处理日期或时间的任何需求，从而为外接程序的不同客户端加强 "一次性写入"。


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>应用场景 A：显示项创建时间或修改时间

如果要在用户界面中显示项创建时间 (**Item.dateTimeCreated**) 或修改时间 (**Item.dateTimeModified**)，请首先使用 **convertToLocalClientTime** 转换这些属性提供的 **Date** 对象以获取采用正确本地时间的字典表示形式。 然后显示字典日期的各个部分。 下面是此应用场景的一个示例：


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

请注意， **convertToLocalClientTime**会考虑 outlook 富客户端和 web 或移动设备上的 outlook 之间的差异：


- 如果 **convertToLocalClientTime** 检测到当前的主机为富客户端，那么该方法使用同一客户端计算机时区（与富客户端用户界面的其余部分保持一致）将 **Date** 表示形式转换为字典表示形式。
    
- 如果**convertToLocalClientTime**检测到当前主机是在 web 或移动设备上的 outlook，则此方法会将 UTC 格式的**日期**表示形式转换为 EAC 时区中的词典格式，与 Outlook 在 web 或移动设备用户界面上的其余部分保持一致。
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>应用场景 B：在新的约会窗体中显示开始日期和结束日期

如果要获取作为输入不同组成部分的以本地时间形式表示的日期时间值，并且希望在约会窗体中将该字典输入值作为开始或结束时间提供，请首先使用 **convertToUtcClientTime** 帮助程序方法将字典值转换为采用 UTC 格式的 **Date** 对象。

在以下示例中，假定 `myLocalDictionaryStartDate` 和 `myLocalDictionaryEndDate` 是从用户获得的采用字典格式的日期时间值。这些值基于本地时间，具体取决于主机应用程序。

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

结果值 `myUTCCorrectStartDate` 和 `myUTCCorrectEndDate` 采用 UTC 格式。 然后将这些 **Date** 对象作为 **Mailbox.displayNewAppointmentForm** 方法的 _Start_ 和 _End_ 参数的自变量来显示新约会窗体。

请注意， **convertToUtcClientTime**会考虑 outlook 富客户端和 web 或移动设备上的 outlook 之间的差异：


- 如果 **convertToUtcClientTime** 检测到当前主机为 Outlook 富客户端，那么该方法只是将字典表示形式转换为 **Date** 对象。 此 **Date** 对象采用 UTC 格式，正如 **displayNewAppointmentForm** 期望的那样。
    
- 如果**convertToUtcClientTime**检测到当前主机是在 web 或移动设备上使用 Outlook，则此方法会将在 EAC 时区中表示的日期和时间值的词典格式转换为**date**对象。 此 **Date** 对象采用 UTC 格式，正如 **displayNewAppointmentForm** 期望的那样。
    

## <a name="see-also"></a>另请参阅

- [部署和安装 Outlook 加载项以进行测试](testing-and-tips.md)
    


