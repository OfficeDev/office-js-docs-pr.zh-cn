---
title: 获取和设置 Outlook 加载项中的定期
description: 本主题介绍如何使用 Office JavaScript API 获取和设置 Outlook 加载项中某个项目的不同定期属性。
ms.date: 08/18/2020
localization_priority: Normal
ms.openlocfilehash: 0b179725677f071fe2ae7baf1c719add5ccd8aa7
ms.sourcegitcommit: e9f23a2857b90a7c17e3152292b548a13a90aa33
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/19/2020
ms.locfileid: "46803742"
---
# <a name="get-and-set-recurrence"></a><span data-ttu-id="2ac11-103">获取和设置定期</span><span class="sxs-lookup"><span data-stu-id="2ac11-103">Get and set recurrence</span></span>

<span data-ttu-id="2ac11-104">有时候，你需要创建和更新某个定期约会，例如团队项目的每周状态会议或每年生日提醒。</span><span class="sxs-lookup"><span data-stu-id="2ac11-104">Sometimes you need to create and update a recurring appointment, such as a weekly status meeting for a team project or a yearly birthday reminder.</span></span> <span data-ttu-id="2ac11-105">您可以使用 Office JavaScript API 来管理外接程序中的约会系列的定期模式。</span><span class="sxs-lookup"><span data-stu-id="2ac11-105">You can use the Office JavaScript API to manage the recurrence patterns of an appointment series in your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="2ac11-106">对此功能的支持是在要求集1.7 中引入的。</span><span class="sxs-lookup"><span data-stu-id="2ac11-106">Support for this feature was introduced in requirement set 1.7.</span></span> <span data-ttu-id="2ac11-107">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="2ac11-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-recurrence-patterns"></a><span data-ttu-id="2ac11-108">可用定期模式</span><span class="sxs-lookup"><span data-stu-id="2ac11-108">Available recurrence patterns</span></span>

<span data-ttu-id="2ac11-109">要配置定期模式，你需要结合使用[定期类型](/javascript/api/outlook/office.mailboxenums.recurrencetype)及其适用的[定期属性](/javascript/api/outlook/office.recurrenceproperties)（如有）。</span><span class="sxs-lookup"><span data-stu-id="2ac11-109">To configure the recurrence pattern, you need to combine the [recurrence type](/javascript/api/outlook/office.mailboxenums.recurrencetype) and its applicable [recurrence properties](/javascript/api/outlook/office.recurrenceproperties) (if any).</span></span>

<span data-ttu-id="2ac11-110">**表 1. 定期类型及其适用的属性**</span><span class="sxs-lookup"><span data-stu-id="2ac11-110">**Table 1. Recurrence types and their applicable properties**</span></span>

|<span data-ttu-id="2ac11-111">定期类型</span><span class="sxs-lookup"><span data-stu-id="2ac11-111">Recurrence type</span></span>|<span data-ttu-id="2ac11-112">有效的定期属性</span><span class="sxs-lookup"><span data-stu-id="2ac11-112">Valid recurrence properties</span></span>|<span data-ttu-id="2ac11-113">用法</span><span class="sxs-lookup"><span data-stu-id="2ac11-113">Usage</span></span>|
|---|---|---|
|`daily`|-&nbsp;[`interval`][interval link]|<span data-ttu-id="2ac11-114">每 *interval* 天进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-114">An appointment occurs every *interval* days.</span></span> <span data-ttu-id="2ac11-115">示例：每 **_2_** 天进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-115">Example: An appointment occurs every **_2_** days.</span></span>|
|`weekday`|<span data-ttu-id="2ac11-116">无。</span><span class="sxs-lookup"><span data-stu-id="2ac11-116">None.</span></span>|<span data-ttu-id="2ac11-117">每个工作日进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-117">An appointment occurs every weekday.</span></span>|
|`monthly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]|<span data-ttu-id="2ac11-118">- 每 *interval* 个月在 *dayOfMonth* 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-118">- An appointment occurs on day *dayOfMonth* every *interval* months.</span></span> <span data-ttu-id="2ac11-119">示例：每 **_4_** 个月在 **_5_** 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-119">Example: An appointment occurs on day **_5_** every **_4_** months.</span></span><br/><br/><span data-ttu-id="2ac11-120">- 每 *interval* 个月在第 *weekNumber* 周的周 *dayOfWeek* 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-120">- An appointment occurs on the *weekNumber* *dayOfWeek* every *interval* months.</span></span> <span data-ttu-id="2ac11-121">示例：每 **_2_** 个月在第**_三_** 周的周**_四_** 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-121">Example: An appointment occurs on the **_third_** **_Thursday_** every **_2_** months.</span></span>|
|`weekly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`days`][days link]|<span data-ttu-id="2ac11-122">每 *interval* 个星期在第 *days* 天进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-122">An appointment occurs on *days* every *interval* weeks.</span></span> <span data-ttu-id="2ac11-123">示例：每 **_2_** 个星期在周**_二_和_四_** 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-123">Example: An appointment occurs on **_Tuesday_ and _Thursday_** every **_2_** weeks.</span></span>|
|`yearly`|-&nbsp;[`interval`][interval link]<br/>-&nbsp;[`dayOfMonth`][dayOfMonth link]<br/>-&nbsp;[`dayOfWeek`][dayOfWeek link]<br/>-&nbsp;[`weekNumber`][weekNumber link]<br/>-&nbsp;[`month`][month link]|<span data-ttu-id="2ac11-124">- 每 *interval* 年在 *month* 月的 *dayOfMonth* 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-124">- An appointment occurs on day *dayOfMonth* of *month* every *interval* years.</span></span> <span data-ttu-id="2ac11-125">示例：每 **_4_** 年在 **_9_** 月 **_7_** 号进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-125">Example: An appointment occurs on day **_7_** of **_September_** every **_4_** years.</span></span><br/><br/><span data-ttu-id="2ac11-126">- 每 *interval* 年在 *month* 月第 *weekNumber* 周的周 *dayOfWeek* 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-126">- An appointment occurs on the *weekNumber* *dayOfWeek* of *month* every *interval* years.</span></span> <span data-ttu-id="2ac11-127">示例：每 **_2_** 年在 **_9_** 月第**_一_** 周的周**_四_** 进行一次约会。</span><span class="sxs-lookup"><span data-stu-id="2ac11-127">Example: An appointment occurs on the **_first_** **_Thursday_** of **_September_** every **_2_** years.</span></span>|

> [!NOTE]
> <span data-ttu-id="2ac11-128">你还可以使用 [`firstDayOfWeek`][firstDayOfWeek link] 属性及 `weekly` 定期类型。</span><span class="sxs-lookup"><span data-stu-id="2ac11-128">You can also use the [`firstDayOfWeek`][firstDayOfWeek link] property with the `weekly` recurrence type.</span></span> <span data-ttu-id="2ac11-129">指定的某一天将从定期对话框中显示的天数列表开始算起。</span><span class="sxs-lookup"><span data-stu-id="2ac11-129">The specified day will start the list of days displayed in the recurrence dialog.</span></span>

## <a name="access-recurrence"></a><span data-ttu-id="2ac11-130">访问定期</span><span class="sxs-lookup"><span data-stu-id="2ac11-130">Access recurrence</span></span>

<span data-ttu-id="2ac11-131">如何访问定期模式以及可对其执行的操作取决于你是约会组织者还是参与者。</span><span class="sxs-lookup"><span data-stu-id="2ac11-131">How you access the recurrence pattern and what you can do with it depends on if you're the appointment organizer or an attendee.</span></span>

<span data-ttu-id="2ac11-132">**表 2. 适用的约会状态**</span><span class="sxs-lookup"><span data-stu-id="2ac11-132">**Table 2. Applicable appointment states**</span></span>

|<span data-ttu-id="2ac11-133">约会状态</span><span class="sxs-lookup"><span data-stu-id="2ac11-133">Appointment state</span></span>|<span data-ttu-id="2ac11-134">约会是否可编辑？</span><span class="sxs-lookup"><span data-stu-id="2ac11-134">Is recurrence editable?</span></span>|<span data-ttu-id="2ac11-135">约会是否可查看？</span><span class="sxs-lookup"><span data-stu-id="2ac11-135">Is recurrence viewable?</span></span>|
|---|---|---|
|<span data-ttu-id="2ac11-136">约会组织者 - 撰写系列</span><span class="sxs-lookup"><span data-stu-id="2ac11-136">Appointment organizer - compose series</span></span>|<span data-ttu-id="2ac11-137">是 ([`setAsync`][setAsync link])</span><span class="sxs-lookup"><span data-stu-id="2ac11-137">Yes ([`setAsync`][setAsync link])</span></span>|<span data-ttu-id="2ac11-138">是 ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="2ac11-138">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="2ac11-139">约会组织者 - 撰写实例</span><span class="sxs-lookup"><span data-stu-id="2ac11-139">Appointment organizer - compose instance</span></span>|<span data-ttu-id="2ac11-140">否（`setAsync` 返回错误）</span><span class="sxs-lookup"><span data-stu-id="2ac11-140">No (`setAsync` returns an error)</span></span>|<span data-ttu-id="2ac11-141">是 ([`getAsync`][getAsync link])</span><span class="sxs-lookup"><span data-stu-id="2ac11-141">Yes ([`getAsync`][getAsync link])</span></span>|
|<span data-ttu-id="2ac11-142">约会参与者 - 读取系列</span><span class="sxs-lookup"><span data-stu-id="2ac11-142">Appointment attendee - read series</span></span>|<span data-ttu-id="2ac11-143">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="2ac11-143">No (`setAsync` not available)</span></span>|<span data-ttu-id="2ac11-144">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2ac11-144">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="2ac11-145">约会参与者 - 读取实例</span><span class="sxs-lookup"><span data-stu-id="2ac11-145">Appointment attendee - read instance</span></span>|<span data-ttu-id="2ac11-146">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="2ac11-146">No (`setAsync` not available)</span></span>|<span data-ttu-id="2ac11-147">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2ac11-147">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="2ac11-148">会议请求 - 读取系列</span><span class="sxs-lookup"><span data-stu-id="2ac11-148">Meeting request - read series</span></span>|<span data-ttu-id="2ac11-149">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="2ac11-149">No (`setAsync` not available)</span></span>|<span data-ttu-id="2ac11-150">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2ac11-150">Yes ([`item.recurrence`][item.recurrence link])</span></span>|
|<span data-ttu-id="2ac11-151">会议请求 - 读取实例</span><span class="sxs-lookup"><span data-stu-id="2ac11-151">Meeting request - read instance</span></span>|<span data-ttu-id="2ac11-152">否（`setAsync` 不可用）</span><span class="sxs-lookup"><span data-stu-id="2ac11-152">No (`setAsync` not available)</span></span>|<span data-ttu-id="2ac11-153">是 ([`item.recurrence`][item.recurrence link])</span><span class="sxs-lookup"><span data-stu-id="2ac11-153">Yes ([`item.recurrence`][item.recurrence link])</span></span>|

## <a name="set-recurrence-as-the-organizer"></a><span data-ttu-id="2ac11-154">以组织者身份设置定期</span><span class="sxs-lookup"><span data-stu-id="2ac11-154">Set recurrence as the organizer</span></span>

<span data-ttu-id="2ac11-155">除了定期模式之外，你还需要确定约会系列的开始和结束日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2ac11-155">Along with the recurrence pattern, you also need to determine the start and end dates and times of your appointment series.</span></span> <span data-ttu-id="2ac11-156">可通过 [`SeriesTime`][SeriesTime link] 对象管理此信息。</span><span class="sxs-lookup"><span data-stu-id="2ac11-156">The [`SeriesTime`][SeriesTime link] object is used to manage that information.</span></span>

<span data-ttu-id="2ac11-157">约会组织者只能在撰写模式下指定约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="2ac11-157">The appointment organizer can specify the recurrence pattern for an appointment series in compose mode only.</span></span> <span data-ttu-id="2ac11-158">在以下示例中，约会系列已设置为在 2019 年 11 月 2 日至 2019 年 12 月 2 日之间的每个周二和周四上午 10:30 至上午 11:00（太平洋标准时间）进行。</span><span class="sxs-lookup"><span data-stu-id="2ac11-158">In the following example, the appointment series is set to occur from 10:30 AM to 11:00 AM PST every Tuesday and Thursday during the period November 2, 2019 to December 2, 2019.</span></span>

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

## <a name="change-recurrence-as-the-organizer"></a><span data-ttu-id="2ac11-159">将重复周期更改为组织者</span><span class="sxs-lookup"><span data-stu-id="2ac11-159">Change recurrence as the organizer</span></span>

<span data-ttu-id="2ac11-160">在下面的示例中，在撰写模式下，约会组织者获取约会系列的定期对象（给定系列或该系列的实例），然后设置新的定期持续时间。</span><span class="sxs-lookup"><span data-stu-id="2ac11-160">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series, then sets a new recurrence duration.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var recurrencePattern = asyncResult.value;
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

## <a name="get-recurrence"></a><span data-ttu-id="2ac11-161">获取定期</span><span class="sxs-lookup"><span data-stu-id="2ac11-161">Get recurrence</span></span>

### <a name="get-recurrence-as-the-organizer"></a><span data-ttu-id="2ac11-162">以组织者身份获取定期</span><span class="sxs-lookup"><span data-stu-id="2ac11-162">Get recurrence as the organizer</span></span>

<span data-ttu-id="2ac11-163">在以下示例中，在撰写模式下，如果存在系列或该系列的某一实例，则约会组织者可以获取约会系列的定期对象。</span><span class="sxs-lookup"><span data-stu-id="2ac11-163">In the following example, in compose mode, the appointment organizer gets the recurrence object of an appointment series given the series or an instance of that series.</span></span>

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

<span data-ttu-id="2ac11-164">以下示例展示了用于检索某个系列定期的 `getAsync` 调用的结果。</span><span class="sxs-lookup"><span data-stu-id="2ac11-164">The following example shows the results of the `getAsync` call that retrieves the recurrence for a series.</span></span>

> [!NOTE]
> <span data-ttu-id="2ac11-165">在此实例中，`seriesTimeObject` 是表示 `recurrence.seriesTime` 属性的 JSON 的占位符。</span><span class="sxs-lookup"><span data-stu-id="2ac11-165">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="2ac11-166">应使用 [`SeriesTime`][SeriesTime link] 方法获取定期日期和时间属性。</span><span class="sxs-lookup"><span data-stu-id="2ac11-166">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-recurrence-as-an-attendee"></a><span data-ttu-id="2ac11-167">以参与者身份获取定期</span><span class="sxs-lookup"><span data-stu-id="2ac11-167">Get recurrence as an attendee</span></span>

<span data-ttu-id="2ac11-168">在以下示例中，如果存在系列、该系列的某一实例或者会议请求，则约会参与者可以获取约会系列的定期对象。</span><span class="sxs-lookup"><span data-stu-id="2ac11-168">In the following example, an appointment attendee can get the recurrence object of an appointment series given the series, an instance of that series, or a meeting request.</span></span>

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

<span data-ttu-id="2ac11-169">以下示例展示了约会系列的 `item.recurrence` 属性值。</span><span class="sxs-lookup"><span data-stu-id="2ac11-169">The following example shows the value of the `item.recurrence` property for an appointment series.</span></span>

> [!NOTE]
> <span data-ttu-id="2ac11-170">在此实例中，`seriesTimeObject` 是表示 `recurrence.seriesTime` 属性的 JSON 的占位符。</span><span class="sxs-lookup"><span data-stu-id="2ac11-170">In this example, `seriesTimeObject` is a placeholder for the JSON representing the `recurrence.seriesTime` property.</span></span> <span data-ttu-id="2ac11-171">应使用 [`SeriesTime`][SeriesTime link] 方法获取定期日期和时间属性。</span><span class="sxs-lookup"><span data-stu-id="2ac11-171">You should use the [`SeriesTime`][SeriesTime link] methods to get the recurrence date and time properties.</span></span>

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

### <a name="get-the-recurrence-details"></a><span data-ttu-id="2ac11-172">获取定期详细信息</span><span class="sxs-lookup"><span data-stu-id="2ac11-172">Get the recurrence details</span></span>

<span data-ttu-id="2ac11-173">检索到定期对象（通过 `getAsync` 回调或通过 `item.recurrence`）之后，可以获取定期的特定属性。</span><span class="sxs-lookup"><span data-stu-id="2ac11-173">After you've retrieved the recurrence object (either from the `getAsync` callback or from `item.recurrence`), you can get specific properties of the recurrence.</span></span> <span data-ttu-id="2ac11-174">例如，可以使用  属性上的[方法][SeriesTime link]`recurrence.seriesTime`获取系列的开始和结束日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2ac11-174">For example, you can get the start and end dates and times of the series by using [methods][SeriesTime link] on the `recurrence.seriesTime` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="2ac11-175">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2ac11-175">See also</span></span>

[<span data-ttu-id="2ac11-176">RecurrenceChanged 事件</span><span class="sxs-lookup"><span data-stu-id="2ac11-176">RecurrenceChanged event</span></span>](/javascript/api/office/office.eventtype)

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
