---
title: 获取和设置 Outlook 加载项中的定期
description: 本主题介绍如何使用 Office JavaScript API 获取和设置 Outlook 加载项中某个项目的不同定期属性。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: cc7160ed8fe82a02ced9c03bab181df57e66bb54
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166039"
---
# <a name="get-and-set-recurrence"></a><span data-ttu-id="143e3-103">获取和设置定期</span><span class="sxs-lookup"><span data-stu-id="143e3-103">Get and set recurrence</span></span>

<span data-ttu-id="143e3-104">有时候，你需要创建和更新某个定期约会，例如团队项目的每周状态会议或每年生日提醒。</span><span class="sxs-lookup"><span data-stu-id="143e3-104">Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder.</span></span> <span data-ttu-id="143e3-105">你可以使用适用于 Office 的 JavaScript API 来管理加载项中约会系列的定期模式。</span><span class="sxs-lookup"><span data-stu-id="143e3-105">You can use the JavaScript API for Office to manage the recurrence patterns of an appointment series in your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="143e3-106">对此功能的支持是在要求集1.7 中引入的。</span><span class="sxs-lookup"><span data-stu-id="143e3-106">Support for this feature was introduced in requirement set 1.7.</span></span> <span data-ttu-id="143e3-107">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="143e3-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-recurrence-patterns"></a><span data-ttu-id="143e3-108">可用定期模式</span><span class="sxs-lookup"><span data-stu-id="143e3-108">Available recurrence patterns</span></span>

<span data-ttu-id="143e3-109">要配置定期模式，你需要结合使用[定期类型](/javascript/api/outlook/office.mailboxenums.recurrencetype)及其适用的[定期属性](/javascript/api/outlook/office.recurrenceproperties)（如有）。</span><span class="sxs-lookup"><span data-stu-id="143e3-109">To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).</span></span>

<span data-ttu-id="143e3-110">**表 1. 定期类型及其适用的属性**</span><span class="sxs-lookup"><span data-stu-id="143e3-110">**Table 1. Recurrence types and their applicable properties**</span></span>

|<span data-ttu-id="143e3-111">定期类型</span><span class="sxs-lookup"><span data-stu-id="143e3-111">Recurrence type</span></span>|<span data-ttu-id="143e3-112">有效的定期属性</span><span class="sxs-lookup"><span data-stu-id="143e3-112">Valid recurrence properties</span></span>|<span data-ttu-id="143e3-113">用法</span><span class="sxs-lookup"><span data-stu-id="143e3-113">Usage</span></span>|
|---|---|---|
|`daily`|- [`interval`][interval link]|<span data-ttu-id="143e3-114">每 *interval* 天进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-114">An appointment occurs every *interval* days.</span></span> <span data-ttu-id="143e3-115">示例：每 **_2_** 天进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-115">Example: An appointment occurs every **_2_** days.</span></span>|
|`weekday`|<span data-ttu-id="143e3-116">无。</span><span class="sxs-lookup"><span data-stu-id="143e3-116">None.</span></span>|<span data-ttu-id="143e3-117">每个工作日进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-117">An appointment occurs every weekday.</span></span>|
|`monthly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]|<span data-ttu-id="143e3-118">- 每 *interval* 个月在 *dayOfMonth* 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-118">- An appointment occurs on day *dayOfMonth* every *interval* months.</span></span> <span data-ttu-id="143e3-119">示例：每 **_4_** 个月在 **_5_** 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-119">Example: An appointment occurs on day **_5_** every **_4_** months.</span></span><br/><br/><span data-ttu-id="143e3-120">- 每 *interval* 个月在第 *weekNumber* 周的周 *dayOfWeek* 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-120">- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months.</span></span> <span data-ttu-id="143e3-121">示例：每 **_2_** 个月在第**_三_** 周的周**_四_** 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-121">Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.</span></span>|
|`weekly`|- [`interval`][interval link]<br/>- [`days`][days link]|<span data-ttu-id="143e3-122">每 *interval* 个星期在第 *days* 天进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-122">An appointment occurs on *days* every *interval* weeks.</span></span> <span data-ttu-id="143e3-123">示例：每 **_2_** 个星期在周**_二_和_四_** 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-123">Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.</span></span>|
|`yearly`|- [`interval`][interval link]<br/>- [`dayOfMonth`][dayOfMonth link]<br/>- [`dayOfWeek`][dayOfWeek link]<br/>- [`weekNumber`][weekNumber link]<br/>- [`month`][month link]|<span data-ttu-id="143e3-124">- 每 *interval* 年在 *month* 月的 *dayOfMonth* 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-124">- An appointment occurs on day *dayOfMonth* of *month* every *interval* years.</span></span> <span data-ttu-id="143e3-125">示例：每 **_4_** 年在 **_9_** 月 **_7_** 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-125">Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.</span></span><br/><br/><span data-ttu-id="143e3-126">- 每 *interval* 年在 *month* 月第 *weekNumber* 周的周 *dayOfWeek* 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-126">- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years.</span></span> <span data-ttu-id="143e3-127">示例：每 **_2_** 年在 **_9_** 月第**_一_** 周的周**_四_** 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="143e3-127">Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.</span></span>|

> [!NOTE]
> <span data-ttu-id="143e3-128">你还可以使用 [`firstDayOfWeek`][firstDayOfWeek link] 属性及 `weekly` 定期类型。</span><span class="sxs-lookup"><span data-stu-id="143e3-128">You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type.</span></span> <span data-ttu-id="143e3-129">指定的某一天将从定期对话框中显示的天数列表开始算起。</span><span class="sxs-lookup"><span data-stu-id="143e3-129">The specified day will start the list of days displayed in the recurrence dialog.</span></span>

## <a name="access-recurrence"></a><span data-ttu-id="143e3-130">访问定期</span><span class="sxs-lookup"><span data-stu-id="143e3-130">Access recurrence</span></span>

<span data-ttu-id="143e3-131">如何访问定期模式以及可对其执行的操作取决于你是约会组织者还是参与者。</span><span class="sxs-lookup"><span data-stu-id="143e3-131">How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.</span></span>

<span data-ttu-id="143e3-132">**表 2. 适用的约会状态**</span><span class="sxs-lookup"><span data-stu-id="143e3-132">**Table 2. Applicable appointment states**</span></span>

|<span data-ttu-id="143e3-133">约会状态</span><span class="sxs-lookup"><span data-stu-id="143e3-133">Appointment state</span></span>|<span data-ttu-id="143e3-134">约会是否可编辑？</span><span class="sxs-lookup"><span data-stu-id="143e3-134">Is recurrence editable?</span></span>|<span data-ttu-id="143e3-135">约会是否可查看？</span><span class="sxs-lookup"><span data-stu-id="143e3-135">Is recurrence viewable?</span></span>|
|---|---|---|
|<span data-ttu-id="143e3-136">约会组织者 - 撰写系列</span><span class="sxs-lookup"><span data-stu-id="143e3-136">Appointment organizer - compose series</span></span>|<span data-ttu-id="143e3-137">是 ([`setAsync`][setAsync link])</span><span class="sxs-lookup"><span data-stu-id="143e3-137">Yes ([`setAsync`][setAsync link])</span></span>|<span data-ttu-id="143e3-138">是 ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="143e3-138">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="143e3-139">约会组织者 - 撰写实例</span><span class="sxs-lookup"><span data-stu-id="143e3-139">Appointment organizer - compose instance</span></span>|<span data-ttu-id="143e3-140">否（`setAsync` 返回错误）</span><span class="sxs-lookup"><span data-stu-id="143e3-140">No (`setAsync` returns an error)</span></span>|<span data-ttu-id="143e3-141">是 ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="143e3-141">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="143e3-142">约会参与者 - 读取系列</span><span class="sxs-lookup"><span data-stu-id="143e3-142">Appointment attendee - read series</span></span>|<span data-ttu-id="143e3-143">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="143e3-143">No (`setAsync` not available)</span></span>|<span data-ttu-id="143e3-144">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="143e3-144">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="143e3-145">约会参与者 - 读取实例</span><span class="sxs-lookup"><span data-stu-id="143e3-145">Appointment attendee - read instance</span></span>|<span data-ttu-id="143e3-146">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="143e3-146">No (`setAsync` not available)</span></span>|<span data-ttu-id="143e3-147">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="143e3-147">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="143e3-148">会议请求 - 读取系列</span><span class="sxs-lookup"><span data-stu-id="143e3-148">Meeting request - read series</span></span>|<span data-ttu-id="143e3-149">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="143e3-149">No (`setAsync` not available)</span></span>|<span data-ttu-id="143e3-150">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="143e3-150">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="143e3-151">会议请求 - 读取实例</span><span class="sxs-lookup"><span data-stu-id="143e3-151">Meeting request - read instance</span></span>|<span data-ttu-id="143e3-152">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="143e3-152">No (`setAsync` not available)</span></span>|<span data-ttu-id="143e3-153">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="143e3-153">Yes ([`item.recurrence`][item.recurrence link])</span></span>|

## <a name="set-recurrence-as-the-organizer"></a><span data-ttu-id="143e3-154">以组织者身份设置定期</span><span class="sxs-lookup"><span data-stu-id="143e3-154">Set recurrence as the organizer</span></span>

<span data-ttu-id="143e3-155">除了定期模式之外，你还需要确定约会系列的开始和结束日期和时间。</span><span class="sxs-lookup"><span data-stu-id="143e3-155">Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series.</span></span> <span data-ttu-id="143e3-156">可通过 [`SeriesTime`][SeriesTime link] 对象管理此信息。</span><span class="sxs-lookup"><span data-stu-id="143e3-156">The [`SeriesTime`][SeriesTime link] object is used to manage that information.</span></span>

<span data-ttu-id="143e3-157">约会组织者只能在撰写模式下指定约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="143e3-157">The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only.</span></span> <span data-ttu-id="143e3-158">在以下示例中，约会系列已设置为在 2019 年 11 月 2 日至 2019 年 12 月 2 日之间的每个周二和周四上午 10:30 至上午 11:00（太平洋标准时间）进行。</span><span class="sxs-lookup"><span data-stu-id="143e3-158">In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.</span></span>

```js
var seriesTimeObject = new Office.SeriesTime();
seriesTimeObject.setStartDate(2019,10,2);
seriesTimeObject.setEndDate(2019,11,2);
seriesTimeObject.setStartTime(10,30);
seriesTimeObject.setDuration(30);

var pattern = {
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

## <a name="get-recurrence"></a><span data-ttu-id="143e3-159">获取定期</span><span class="sxs-lookup"><span data-stu-id="143e3-159">Get recurrence</span></span>

### <a name="get-recurrence-as-the-organizer"></a><span data-ttu-id="143e3-160">以组织者身份获取定期</span><span class="sxs-lookup"><span data-stu-id="143e3-160">Get recurrence as the organizer</span></span>

<span data-ttu-id="143e3-161">在以下示例中，在撰写模式下，如果存在系列或该系列的某一实例，则约会组织者可以获取约会系列的定期对象。</span><span class="sxs-lookup"><span data-stu-id="143e3-161">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult){
    var context = asyncResult.context;
    var recurrence = asyncResult.value;

    if (recurrence == null) {
        console.log("Non-recurring meeting");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

<span data-ttu-id="143e3-162">以下示例展示了用于检索某个系列定期的 `getAsync` 调用的结果。</span><span class="sxs-lookup"><span data-stu-id="143e3-162">The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.</span></span>

> [!NOTE]
> <span data-ttu-id="143e3-163">在此实例中，`seriesTimeObject` 是表示 `recurrence.seriesTime` 属性的 JSON 的占位符。</span><span class="sxs-lookup"><span data-stu-id="143e3-163">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="143e3-164">应使用 [`SeriesTime`][SeriesTime link] 方法获取定期日期和时间属性。</span><span class="sxs-lookup"><span data-stu-id="143e3-164">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-recurrence-as-an-attendee"></a><span data-ttu-id="143e3-165">以参与者身份获取定期</span><span class="sxs-lookup"><span data-stu-id="143e3-165">Get recurrence as an attendee</span></span>

<span data-ttu-id="143e3-166">在以下示例中，如果存在系列、该系列的某一实例或者会议请求，则约会参与者可以获取约会系列的定期对象。</span><span class="sxs-lookup"><span data-stu-id="143e3-166">In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.</span></span>

```js
outputRecurrence(Office.context.mailbox.item);

function outputRecurrence(item) {
    var recurrence = item.recurrence;
    var seriesId = item.seriesId;

    if (recurrence == null) {
        console.log("Non-recurring item");
    } else {
        console.log(JSON.stringify(recurrence));
    }
}
```

<span data-ttu-id="143e3-167">以下示例展示了约会系列的 `item.recurrence` 属性值。</span><span class="sxs-lookup"><span data-stu-id="143e3-167">The following example shows the value of the `item.recurrence` property for an appointment series.</span></span>

> [!NOTE]
> <span data-ttu-id="143e3-168">在此实例中，`seriesTimeObject` 是表示 `recurrence.seriesTime` 属性的 JSON 的占位符。</span><span class="sxs-lookup"><span data-stu-id="143e3-168">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="143e3-169">应使用 [`SeriesTime`][SeriesTime link] 方法获取定期日期和时间属性。</span><span class="sxs-lookup"><span data-stu-id="143e3-169">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-the-recurrence-details"></a><span data-ttu-id="143e3-170">获取定期详细信息</span><span class="sxs-lookup"><span data-stu-id="143e3-170">Get the recurrence details</span></span>

<span data-ttu-id="143e3-171">检索到定期对象（通过 `getAsync` 回调或通过 `item.recurrence`）之后，可以获取定期的特定属性。</span><span class="sxs-lookup"><span data-stu-id="143e3-171">After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence.</span></span> <span data-ttu-id="143e3-172">例如，可以使用  属性上的[方法][SeriesTime link]`recurrence.seriesTime`获取系列的开始和结束日期和时间。</span><span class="sxs-lookup"><span data-stu-id="143e3-172">For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.</span></span>

```js
// Get series date and time info
var seriesTime = recurrence.seriesTime;
var startTime = recurrence.seriesTime.getStartTime();
var endTime = recurrence.seriesTime.getEndTime();
var startDate = recurrence.seriesTime.getStartDate();
var endDate = recurrence.seriesTime.getEndDate();
var duration = recurrence.seriesTime.getDuration();

// Get series time zone
var timeZone = recurrence.recurrenceTimeZone;

// Get recurrence properties
var recurrenceProperties = recurrence.recurrenceProperties;

// Get recurrence type
var recurrenceType = recurrence.recurrenceType;
```

## <a name="see-also"></a><span data-ttu-id="143e3-173">另请参阅</span><span class="sxs-lookup"><span data-stu-id="143e3-173">See also</span></span>

[<span data-ttu-id="143e3-174">RecurrenceChanged 事件</span><span class="sxs-lookup"><span data-stu-id="143e3-174">RecurrenceChanged event</span></span>](/javascript/api/office/office.eventtype)

[getAsync link]: /javascript/api/outlook/office.recurrence#getasync-options--callback-
[item.recurrence link]: ../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties
[setAsync link]: /javascript/api/outlook/office.recurrence#setasync-recurrencepattern--options--callback-

[dayOfMonth link]: /javascript/api/outlook/office.recurrenceproperties#dayofmonth
[dayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#dayofweek
[days link]: /javascript/api/outlook/office.recurrenceproperties#days
[firstDayOfWeek link]: /javascript/api/outlook/office.recurrenceproperties#firstdayofweek
[interval link]: /javascript/api/outlook/office.recurrenceproperties#interval
[month link]: /javascript/api/outlook/office.recurrenceproperties#month
[weekNumber link]: /javascript/api/outlook/office.recurrenceproperties#weeknumber

[SeriesTime link]: /javascript/api/outlook/office.seriestime
