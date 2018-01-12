
# <a name="configuring-and-editing-labsjs-labs-for-office-mix"></a>配置和编辑 Office Mix 的 LabsJS 实验室



Office Mix 提供用于获取和设置实验室配置的 office.js 方法。配置向 Office Mix 指示您创建的实验室类型，以及实验室将发回的数据类型。此信息用于收集分析数据并将其可视化。

## <a name="getting-the-lab-editor"></a>获取实验室编辑器

实验室编辑器 [Labs.LabEditor](../../../reference/office-mix/labs.labeditor.md) 对象允许您编辑实验室并获取和设置您的实验室配置。当您编辑完实验室之后，必须调用 **Done** 方法。但是，调用 **Done** 方法并非必需的，除非您尝试使用或运行您正在编辑的实验室。请注意，一次只能打开实验室的一个实例。

以下代码显示如何获取实验室编辑器。




```js
Labs.editLab((err, labEditor) => {
    if (err) {
        handleError();
        return;
    }
    _labEditor = labEditor;
});
```

使用 **Labs.LabEditor** 上的 **getConfiguration** 和 [setConfiguration](../../../reference/office-mix/labs.labeditor.md) 方法存储指定实验室的配置。配置 ([Labs.Core.IConfiguration](../../../reference/office-mix/labs.core.iconfiguration.md)) 向 Office Mix 指明实验室将收集和处理哪些数据。配置包含关于实验室的常规信息，包括名称、版本和其他配置选项。配置最重要的部分是实验室组件的定义。

以下代码演示如何设置和获取配置。要设置配置，只需创建配置对象，然后调用  **setConfiguration** 方法。要检索配置，您可以对实验室编辑器对象调用 **getConfiguration** 方法。




```js

///////  Set the configuration /////

var activityComponent: Labs.Components.IActivityComponent = {
    type: Labs.Components.ActivityComponentType,
    name: uri,
    values: {},
    data: {
        uri: uri
    },
    secure: false
};
var configuration = {
    appVersion: { major: 1, minor: 1 },
    components: [activityComponent],
    name: configurationName,
    timeline: null,
    analytics: null
};
this._labEditor.setConfiguration(configuration, (err, unused) => { })

```




```js

///////  Get the configuration  //////

labEditor.getConfiguration((err, configuration) => {
});
```


## <a name="closing-the-editor"></a>关闭编辑器

要关闭编辑器，请在编辑完实验室之后对编辑器调用  **Done** 方法。请注意，您无法同时使用和编辑实验室。但是，调用 **Done** 之后，您可以编辑或运行实验室。


## <a name="interacting-with-a-lab"></a>与实验室交互

设置实验室配置后，即表示您已准备好开始与实验室交互。当实验室在 PowerPoint 内运行时，会对交互进行模拟。但是，当实验室在 Office Mix 课程播放器内运行时，数据将存储在 Office Mix 数据库中并在分析时使用。


### <a name="getting-the-lab-instance"></a>获取实验室实例

您可使用 [Labs.LabInstance](../../../reference/office-mix/labs.labinstance.md) 对象与实验室交互，该对象是为当前用户配置的实验室的实例。要运行（或使用）实验室，请调用 [Labs.takeLab](../../../reference/office-mix/labs.takelab.md) 函数。


```js
Labs.takeLab((err, labInstance) => {
    this._labInstance = labInstance;
    var activityComponentInstance = <Labs.Components.ActivityComponentInstance> this._labInstance.components[0];
    // populate the UI based on the instance    
});
```

实例对象包含一系列组件实例（ [Labs.ComponentInstanceBase](../../../reference/office-mix/labs.componentinstancebase.md)、 [Labs.ComponentInstance](../../../reference/office-mix/labs.componentinstance.md)），这些实例会映射到您在配置中指定的组件。实际上，实例就是配置的转换版本，用于将服务器端 ID 附加到实例对象，以及在必要时对用户隐藏某些字段（例如提示、答案等）。


### <a name="managing-state"></a>管理状态

状态是与运行指定实验室的用户相关的临时存储。您可以使用存储保存实验室后续调用之间的信息。例如，编程实验室可以存储用户当前进行的工作。

要 **set**状态，请使用以下代码。




```js
labInstance.setState(this._labState(), (err, unused) => { 
    // If no error, state has successfully been stored by the host.
});
```

要 **get**状态，请使用以下代码。




```js
labInstance.getState((err, state) => {
    // If no error, the state parameter contains the set state.
});
```


## <a name="component-instances-and-results"></a>组件实例和结果

接下来将概述如何实施四种组件类型的实例，并提供了组件方法的简短示例。 

但是，您需要首先熟悉使用组件实例的两个核心概念。第一个概念是 **尝试** 和 **值** 的概念。

 **尝试**

尝试是指用户尝试完成组件实例。例如，在多选题中，当用户开始解决问题时，尝试开始；当分配最终分数时，尝试结束。然后 Office Mix 分析将汇总问题的结果。


 >**注意**：可对除 **DynamicComponent** 类型以外的所有组件类型进行尝试。

您可以使用  **getAttempts** 方法检索与指定组件实例相关的所有尝试的结果。检索结果后，用户可以使用 **resume** 方法重试现有尝试之一，或者使用 **createAttempt** 方法创建新的尝试。以下示例显示了此过程。




```js
var attemptsDeferred = $.Deferred();
activityComponentInstance.getAttempts(createCallback(attemptsDeferred));
var attemptP = attemptsDeferred.promise().then((attempts) => {
    var currentAttemptDeferred = $.Deferred();
    if (attempts.length > 0) {
        currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
    } else {
        activityComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
    }
    return currentAttemptDeferred.then((currentAttempt: Labs.Components.ActivityComponentAttempt) => {
        var resumeDeferred = $.Deferred();
        currentAttempt.resume(createCallback(resumeDeferred));
        return resumeDeferred.promise().then(() => {
            return currentAttempt;
        });
    });
});
```

 **值**

组件实例包含映射到值数组的键的字典。您可以使用此数组存储您希望与组件关联的提示、反馈或任何其他值集。组件实例使用  **getValues** 方法提供对这些值的访问权限。

例如，查询提示值会导致分析标记用户使用了提示。每次尝试时都会跟踪值。

以下代码示例说明了如何查询提示。




```js
// Take a hint.
var hints = attempt.getValues("hints");
hints[0].getValue((err, hint) => {
    // If no error, hint param will contain the hint data.
});
```


### <a name="activitycomponentinstance"></a>ActivityComponentInstance


使用  **ActivityComponentInstace** 对象跟踪用户与活动组件的交互。该类提供了 **complete** 方法，以指示用户已完成与活动的交互。方法可以指示用户已完成分配的任务、读取或与活动相关的任何其他端点。以下代码显示了如何使用 **complete** 方法。


```js
attempt.complete((err, unused) => { 
    // Called after the host has stored the completion.
});
```


### <a name="choicecomponentinstance"></a>ChoiceComponentInstance


使用  **ChoiceComponentInstance** 对象跟踪用户与选择组件的交互。选择组件是将向用户显示他们需要从中选择的选项列表的问题。这些选项可能是，也可能不是正确答案。该类提供两个主要方法： **getSubmissions** 和 **submit**。 **getSubmissions** 方法允许您检索之前存储的提交； **submit** 方法允许存储新的提交。以下代码示例说明了如何使用这两个方法。


```js
///  using getSubmission method  ///
var submissions = this._attempt.getSubmissions();
```


```js
///  using submit method  ///
this._attempt.submit(
    new Labs.Components.ChoiceComponentAnswer(submission), 
    new Labs.Components.ChoiceComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="inputcomponentinstance"></a>InputComponentInstance


使用  **InputComponentInstance** 对象跟踪用户与输入组件的交互。该类提供两个主要方法： **getSubmission** 和 **submit**。 **getSubmissions** 方法允许您检索之前存储的提交； **submit** 方法允许您存储新的提交。以下代码段说明了如何使用 **getSubmissions** 方法。


```js
var submissions = this._attempt.getSubmissions();
```

使用  **submit** 方法时，请注意 **InputComponentAnswer** 对象代表提交的答案， **InputComponentResult** 对象包含结果。返回值是包含答案、结果以及指示结果何时提交的时间戳的 **InputComponentSubmission** 对象。




```js
this._attempt.submit(
    new Labs.Components.InputComponentAnswer(submission), 
    new Labs.Components.InputComponentResult(correct, complete), 
    (err, submission) => {
        // Called after the server has processed the submission.
    });
```


### <a name="dynamiccomponentinstance"></a>DynamicComponentInstance


使用  **DynamicComponentInstance** 对象跟踪用户与动态组件的交互。该类中的主要方法包括 **getComponents**、 **createComponent** 和 **close**。

**getComponents** 方法允许你检索以前创建的组件实例列表，如以下示例中所示。




```js
dynamicComponentInstance.getComponents((err, components) => {
    // Upon success, components contains a list of previously created component instances.
});
```

**createComponent** 方法可构建一个新组件并返回该组件实例，如以下示例中所示。




```js
var inputComponentHints = [];
for (var i = 0; i < data.hints.length; i++) {
    inputComponentHints.push({
        isHint: true,
        value: data.hints[i]        
    });
}
var inputComponent = {
    maxScore: 1,
    timeLimit: 0,
    hasAnswer: true,
    answer: data.answerData.solution,
    type: Labs.Components.InputComponentType,
    name: data.name,
    values: { hints: inputComponentHints },
    secure: false
};
var currentAttemptDeferred = $.Deferred();
var dynamicComponent = labInstance.components[0];
dynamicComponent.createComponent(inputComponent, function(err, inputComponentInstance) {
    // Create will return the instance for the specified component.
})
```

使用  **close** 方法指示您已使用动态组件创建新组件。请注意，您还可以使用 **isClosed** Boolean 方法测试动态组件实例是否已关闭。以下代码显示了如何使用 **close** 方法。




```js
dynamicComponentInstance.close((err, unused) => {
    // Called after the server has processed the close attempt.
});
```


## <a name="additional-resources"></a>其他资源



- [Office Mix 外接程序](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [演练︰创建第一个 Office Mix 实验室](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
