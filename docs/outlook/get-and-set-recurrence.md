---
title: 获取和设置 Outlook 加载项中的定期
description: 本主题介绍如何使用 Office JavaScript API 获取和设置 Outlook 加载项中某个项目的不同定期属性。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: f0fbafcb761a74e5a28294c25b480f4cb35a92fa
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889355"
---
# <a name="get-and-set-recurrence"></a>获取和设置定期

有时候，你需要创建和更新某个定期约会，例如团队项目的每周状态会议或每年生日提醒。 可以使用 Office JavaScript API 管理外接程序中约会系列的重复模式。

> [!NOTE]
> [要求集 1.7](/javascript/api/requirement-sets/outlook/requirement-set-1.7/outlook-requirement-set-1.7) 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="available-recurrence-patterns"></a>可用定期模式

要配置定期模式，你需要结合使用[定期类型](/javascript/api/outlook/office.mailboxenums.recurrencetype)及其适用的[定期属性](/javascript/api/outlook/office.recurrenceproperties)（如有）。

**表 1. 定期类型及其适用的属性**

|定期类型|有效的定期属性|用法|
|---|---|---|
|`daily`|-&nbsp;[`interval`][interval link]|每 *interval* 天进行一次约会。 示例：每 **_2_** 天进行一次约会。|
|`weekday`|无。|每个工作日进行一次约会。|
|`monthly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]|- 每 *interval* 个月在 *dayOfMonth* 号进行一次约会。 示例：每 **_4_** 个月在 **_5_** 号进行一次约会。<br/><br/>- 每 *interval* 个月在第 *weekNumber* 周的周 *dayOfWeek* 进行一次约会。 示例：每 **_2_** 个月在第 **_三_** 周的周 **_四_** 进行一次约会。|
|`weekly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`days`][days link]|每 *interval* 个星期在第 *days* 天进行一次约会。 示例：每 **_2_** 个星期在周 **_二_ 和 _四_** 进行一次约会。|
|`yearly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]<br/>-&nbsp;[`month`][month link]|- 每 *interval* 年在 *month* 月的 *dayOfMonth* 号进行一次约会。 示例：每 **_4_** 年在 **_9_** 月 **_7_** 号进行一次约会。<br/><br/>- 每 *interval* 年在 *month* 月第 *weekNumber* 周的周 *dayOfWeek* 进行一次约会。 示例：每 **_2_** 年在 **_9_** 月第 **_一_** 周的周 **_四_** 进行一次约会。|

> [!NOTE]
> 你还可以使用 [`firstDayOfWeek`][firstDayOfWeek link] 属性及 `weekly` 定期类型。 指定的某一天将从定期对话框中显示的天数列表开始算起。

## <a name="access-recurrence"></a>访问定期

如何访问定期模式以及可对其执行的操作取决于你是约会组织者还是参与者。

**表 2. 适用的约会状态**

|约会状态|约会是否可编辑？|约会是否可查看？|
|---|---|---|
|约会组织者 - 撰写系列|是 ([`setAsync`][setAsync link]) |是 ([`getAsync`][getAsync link]) |
|约会组织者 - 撰写实例|否（`setAsync` 返回错误）|是 ([`getAsync`][getAsync link]) |
|约会参与者 - 读取系列|否（`setAsync` 不可用）|是 ([`item.recurrence`][item.recurrence link]) |
|约会参与者 - 读取实例|否（`setAsync` 不可用）|是 ([`item.recurrence`][item.recurrence link]) |
|会议请求 - 读取系列|否（`setAsync` 不可用）|是 ([`item.recurrence`][item.recurrence link]) |
|会议请求 - 读取实例|否（`setAsync` 不可用）|是 ([`item.recurrence`][item.recurrence link]) |

## <a name="set-recurrence-as-the-organizer"></a>以组织者身份设置定期

除了定期模式之外，你还需要确定约会系列的开始和结束日期和时间。 可通过 [`SeriesTime`][SeriesTime link] 对象管理此信息。

约会组织者只能在撰写模式下指定约会的定期模式。 在以下示例中，约会系列已设置为在 2019 年 11 月 2 日至 2019 年 12 月 2 日之间的每个周二和周四上午 10:30 至上午 11:00（太平洋标准时间）进行。

```js
const seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

const pattern = {
    "seriesTime": seriesTimeObject,
    "recurrenceType": "weekly",
    "recurrenceProperties": {"interval": 1, "days": ["tue", "thu"]},
    "recurrenceTimeZone": {"name": "Pacific Standard Time"}};

Office.context.mailbox.item.recurrence.setAsync(pattern, callback);

function callback(asyncResult)
{
    console.log(JSON.stringify(asyncResult));
}
```

## <a name="change-recurrence-as-the-organizer"></a>更改作为组织者的重复性

在下面的示例中，在撰写模式下，约会组织者获取给定序列或该系列实例的约会系列的定期对象，然后设置新的定期持续时间。

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  const recurrencePattern = asyncResult.value;
  recurrencePattern.seriesTime.setDuration(60);
  Office.context.mailbox.item.recurrence.setAsync(recurrencePattern, (asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.log("failed");
      return;
    }

    console.log("success");
  });
}
```

## <a name="get-recurrence"></a>获取定期

### <a name="get-recurrence-as-the-organizer"></a>以组织者身份获取定期

在以下示例中，在撰写模式下，如果存在系列或该系列的某一实例，则约会组织者可以获取约会系列的定期对象。

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    const context = asyncResult.context;
    const recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

以下示例展示了用于检索某个系列定期的 `getAsync` 调用的结果。

> [!NOTE]
> 在此实例中，`seriesTimeObject` 是表示 `recurrence.seriesTime` 属性的 JSON 的占位符。 应使用 [`SeriesTime`][SeriesTime link] 方法获取定期日期和时间属性。

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-recurrence-as-an-attendee"></a>以参与者身份获取定期

在以下示例中，如果存在系列、该系列的某一实例或者会议请求，则约会参与者可以获取约会系列的定期对象。

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    const recurrence = item.recurrence;
    const seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

以下示例展示了约会系列的 `item.recurrence` 属性值。

> [!NOTE]
> 在此实例中，`seriesTimeObject` 是表示 `recurrence.seriesTime` 属性的 JSON 的占位符。 应使用 [`SeriesTime`][SeriesTime link] 方法获取定期日期和时间属性。

```json
{
    "recurrenceType": "weekly",
    "recurrenceProperties": {
        "interval": 1,
        "days": ["tue","thu"],
        "firstDayOfWeek": "sun"},
    "seriesTime": {seriesTimeObject},
    "recurrenceTimeZone": {
        "name": "Pacific Standard Time",
        "offset": -480}}
```

### <a name="get-the-recurrence-details"></a>获取定期详细信息

检索到定期对象（通过 `getAsync` 回调或通过 `item.recurrence`）之后，可以获取定期的特定属性。 例如，可以使用  属性上的[方法][SeriesTime link]`recurrence.seriesTime`获取系列的开始和结束日期和时间。

```js
// Get series date and time info
const seriesTime = recurrence.seriesTime;
const startTime = recurrence.seriesTime.getStartTime();
const endTime = recurrence.seriesTime.getEndTime();
const startDate = recurrence.seriesTime.getStartDate();
const endDate = recurrence.seriesTime.getEndDate();
const duration = recurrence.seriesTime.getDuration();

// Get series time zone
const timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
const recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
const recurrenceType = recurrence.recurrenceType;
```

## <a name="see-also"></a>另请参阅

- [RecurrenceChanged 事件](/javascript/api/office/office.eventtype)
- [Recurrence 对象](/javascript/api/outlook/office.recurrence)
- [SeriesTime 对象](/javascript/api/outlook/office.seriestime)

[getAsync link]: /javascript/api/outlook/office.recurrence#getAsync_options__callback_
[item.recurrence link]: /javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setAsync_recurrencePattern__options__callback_

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayOfMonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayOfWeek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstDayOfWeek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weekNumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
