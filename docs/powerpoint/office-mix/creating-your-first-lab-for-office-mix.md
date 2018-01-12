
# <a name="walkthrough-creating-your-first-lab-for-office-mix"></a>演练：为 Office Mix 创建第一个实验室
使用分步演练构建您的第一个 LabsJS 实验室。



在本演练中，你将从零开始创建简单的 LabsJS 实验室。你的实验室将具有简单的正/误判断测验，其中仅提供一个问题。 

你不是从 Visual Studio 项目模板开始，而是仅从三个空文件开始 – 这显示了实验室是多么简单： 


- TrueFalse.html (html5)
    
- TrueFalse.js
    
- TrueFalse.css
    
你可以使用任何你想要用于编辑这些文件的代码编辑器，因为我们未启动 Visual Studio 模板。事实上，HTML 文件很简单，如果你愿意，你可以仅从教程文件中复制/粘贴 HTML 标记。但是，请注意，其必须是 HTML5，因此请确保你的 doctype 声明是 `<!DOCTYPE html>`CSS 文件属于可选项。所有复杂的工作都在 JavaScript (.js) 文件、TrueFalse.js 中完成。演练将覆盖四种主要实验室功能：

- 设置（连接到主机）
    
- 模式更改（在编辑模式和查看模式之间）
    
- 编辑实验室
    
- 提取（或运行）实验室
    

 **注意**  
 ---
 文件 labhost.html 在 Web 服务器上运行并提供用于实验室开发和测试的托管环境。这大大简化了实验室环境。有关如何设置开发环境的信息，请参阅[适用于 Office Mix 的 LabsJS 入门](get-started-with-labsjs-for-office-mix.md)<br/><br/>

最后，你可以在使用此 SDK 分发的文件中查看完成的 JavaScript 文件 (TrueFalse.js)。接下来是编码过程的演练。

## <a name="connecting-to-the-lab-host"></a>连接到实验室主机

此环境中的实验室能够与我们的实验室主机（用于开发和测试）或 Office.js 主机提供的默认运行时主机一起运行。然后启动函数将使用简单的 if/else 表达式测试哪些托管上下文适用。


```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```

**PostMessageLabHost** 对象在 labhost.html 开发环境中运行，而在生产中，实验室使用 **OfficeJSLabHost** 在 PowerPoint/Office Mix 中运行。

接下来，创建一个帮助程序方法以创建回调，其任务是解决或拒绝你传递的 jQuery 延迟对象。使用此方法 **createCallback** 从 jQuery 承诺进入 labs.js 定义的回调。




```js
function createCallback(deferred) {
    return function (err, data) {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```

我们还会创建一个帮助程序方法来检索给定的问题和答案的实验室配置。




```js
function getConfiguration(question, answer) {
    var choiceComponent = {
        name: question,
        type: Labs.Components.ChoiceComponentType,
        timeLimit: 0,
        maxAttempts: 1,
        choices: [
            { id: "0", name: "True", value: "True" },
            { id: "1", name: "False", value: "False" }],
        maxScore: 1,
        hasAnswer: true,
        answer: answer ? "0" : "1",
        values: null,
        secure: false,
        data: null
    };

    return {
        appVersion: { major: 0, minor: 1 },
        components: [choiceComponent],
        name: question,
        timeline: null,
        analytics: null
    };
}
```


## <a name="mode-changes"></a>模式更改

实验室始终为两种状态或模式之一：**view** 和 **edit**。因此，我们需要一种方法来捕获和保留测验的状态和行为；我们将创建一个类用于此目的。


```js
var TrueFalseQuiz = (function () {
    /**
     * Constructor - takes in the starting mode.
     */
    function TrueFalseQuiz(mode) {
        var self = this;        
        self._modeSwitchP = $.when();
        self._labInstance = null;
        self._labEditor = null;        
      /**
       * Listen for mode changed events and 
       * then switch accordingly. Also set the initial mode state.
       */
        Labs.on(Labs.Core.EventTypes.ModeChanged, function (modeChangedEvent) {
            self.switchUserMode(Labs.Core.LabMode[modeChangedEvent.mode]);
        });
        this.switchUserMode(mode);        
    }
```

此外，我们还提供了一个帮助程序方法，其任务是根据测验问题的答案（即“提交”）是正确还是错误来更新测验的 UI。




```js
    TrueFalseQuiz.prototype._showResults = function(correct) {
        $("#submit-button").removeClass("btn-default");
        $("#submit-button").addClass(correct ? "btn-success" : "btn-danger");
        $("#submit-button").text(correct ? "Correct!" : "Incorrect");

        $("#submit-button").prop("disabled", true);
        $("input:radio[name='quizAnswers']").prop("disabled", true);
    };
```

我们还需要一个函数用于在编辑和查看模式之间切换。




```js
TrueFalseQuiz.prototype.switchUserMode = function (mode) {
        var self = this;

        // Wait for any previous mode switch to complete before performing the new one
        self._modeSwitchP = self._modeSwitchP.then(function () {
            var switchedStateDeferred = $.Deferred();

            // Clean up any variables associated with the previous mode.
            if (self._labInstance) {
                $("#quiz-view-form").off("submit");
                self._labInstance.done(createCallback(switchedStateDeferred));
            } else if (self._labEditor) {
                self._unbindFromEditUpdates();
                self._labEditor.done(createCallback(switchedStateDeferred));
            } else {
                switchedStateDeferred.resolve();
            }

            // After the cleanup occurs, switch to the new mode.
            return switchedStateDeferred.promise().then(function () {
                self._labEditor = null;
                self._labInstance = null;

                if (mode === Labs.Core.LabMode.Edit) {
                    return self._switchToEditMode();
                } else {
                    return self._switchToViewMode();
                }
            });
        });

        // Display an error if it occurs.
        self._modeSwitchP.fail(function (error) {
            // ... error handling ...
        });
    };
```

我们的下一个函数将根据我们从 UI 接收的变更事件来更新测验的配置。




```js
    TrueFalseQuiz.prototype._updateConfigurationFromUI = function () {
        var question = $("#question-edit").val();
        var answerIsTrue = $("input:radio[name='answerValue']:checked").val() === "true";

        this._updateConfiguration(question, answerIsTrue, true, function (err) {
            if (err) {
                // show error
            }
        });
    };
```

接下来，我们将根据指定的问题和答案更新存储在服务器上的实验室配置数据。




```js
    TrueFalseQuiz.prototype._updateConfiguration = function (question, answer, serialize, callback) {
        var configuration = getConfiguration(question, answer);

        if (serialize) {
            this._labEditor.setConfiguration(configuration, callback);
        } else {
            callback(null, null);
        }
    };
```

接下来，我们有一个函数，它会将在编辑模式下在实验室中所做的更新绑定到我们所做的配置更改，随后是用于与以前绑定的更改处理程序取消绑定的代码。




```js
    TrueFalseQuiz.prototype._bindToEditUpdates = function () {
        var self = this;

        // Listen for the question changing
        $("#question-edit").on("input propertychange paste", function () {
            self._updateConfigurationFromUI();
        });

        $('input[name="answerValue"]').on("change", function (e) {
            self._updateConfigurationFromUI();
        });
    };
```




```js
    TrueFalseQuiz.prototype._unbindFromEditUpdates = function () {
        $("#question-edit").off("input propertychange paste");
        $('input[name="answerValue"]').off("change");
    };
```

下面是节的关键部分，即用于在查看和编辑模式之间来回切换的方法。我们先从查看模式切换到编辑模式。




```js
    TrueFalseQuiz.prototype._switchToEditMode = function () {
        var self = this;
        var editLabDeferred = $.Deferred();

        // Make the Labs.js API call to edit the lab.
        Labs.editLab(createCallback(editLabDeferred));

        return editLabDeferred.promise().then(function (labEditor) {            
            self._labEditor = labEditor;

            // Retrieve any existing configuration from the lab editor.
            var configurationDeferred = $.Deferred();
            labEditor.getConfiguration(createCallback(configurationDeferred));

            return configurationDeferred.promise().then(function (configuration) {
                var configurationReadyDeferred = $.Deferred();

                // Get the question and answer values if they exist. 
                //Otherwise use the defaults.
                var question = configuration !== null ? configuration.components[0].name : "";
                var answerIsTrue = configuration !== null ? configuration.components[0].answer === "0" : true;

                // Update the lab configuration based on the question and answer.
                self._updateConfiguration(
                    question,
                    answerIsTrue,
                    configuration === null,
                    createCallback(configurationReadyDeferred));

                // Update the UI based on the question and answer.
                $("#question-edit").val(question);
                $('input[name="answerValue"][value="' + answerIsTrue + '"]').prop('checked', true);

                // Bind to changes.
                self._bindToEditUpdates();

                // Flip over the UI.
                $("#quiz-editor").removeClass("hidden");
                $("#quiz-view").addClass("hidden");

                return configurationReadyDeferred.promise();
            });
        });
    };
```

现在，从编辑模式切换到查看模式。




```js
    TrueFalseQuiz.prototype._switchToViewMode = function () {
        var self = this;
        var takeLabDeferred = $.Deferred();

        // Call the labs.js API to start taking the lab.
        Labs.takeLab(createCallback(takeLabDeferred));

        return takeLabDeferred.promise().then(function (labInstance) {
            self._labInstance = labInstance;

            // Get the choice component instance that will be generated
            // from the choice component we saved when editing the lab.
            var choiceComponentInstance = self._labInstance.components[0];

            // Get the attempts associated with that choice component.
            var attemptsDeferred = $.Deferred();
            choiceComponentInstance.getAttempts(createCallback(attemptsDeferred));
            var attemptP = attemptsDeferred.promise().then(function (attempts) {
                // See if we already had started an attempt against 
                // the problem. If not create one.
                var currentAttemptDeferred = $.Deferred();
                if (attempts.length > 0) {
                    currentAttemptDeferred.resolve(attempts[attempts.length - 1]);
                } else {
                    choiceComponentInstance.createAttempt(createCallback(currentAttemptDeferred));
                }

                return currentAttemptDeferred.then(function (currentAttempt) {
                    var resumeDeferred = $.Deferred();

                    // After we have the attempt, mark that we are resuming
                    // it as well. This will note the resumption time
                    // in the lab activity log.
                    currentAttempt.resume(createCallback(resumeDeferred));
                    return resumeDeferred.promise().then(function () {
                        return currentAttempt;
                    });
                });
            });

            return attemptP.promise().then(function (attempt) {
                // Store off the latest attempt for later use.
                self._currentAttempt = attempt;

                // Update the question field of the view UI.
                $("#question-view").text(choiceComponentInstance.component.name);

                // Determine whether the quiz has already been taken
                // and update the UI accordingly.
                var submissions = attempt.getSubmissions();
                if (submissions.length > 0) {
                    var correctAttempt = submissions[submissions.length - 1].result.score === 1;
                    var submissionValue = submissions[submissions.length - 1].answer.answer === "0";
                    $('input[name="quizAnswers"][value="' + submissionValue + '"]').prop('checked', true);
                    self._showResults(correctAttempt);
                } else {
                    $("#submit-button").removeClass("btn-success btn-danger"    );
                    $("#submit-button").addClass("btn-default");
                    $("#submit-button").text("Submit");
                    $("#submit-button").prop("disabled", false);
                    $("input:radio[name='quizAnswers']").prop("disabled", false);
                }                

                // Hook up the form submit button and then
                // grade the attempt when it is selected.
                $("#quiz-view-form").on("submit", function (e) {
                    e.preventDefault();
                    
                    // Get the checked value and see whether the choice
                    // was true or false - map back to our choice fields.
                    var submission = $("input:radio[name='quizAnswers']:checked").val() === "true" ? "0" : "1";

                    // Grade against the stored answer.
                    var correct = choiceComponentInstance.component.answer === submission;

                    // Submit the attempt with the labs.js API.
                    attempt.submit(
                        new Labs.Components.ChoiceComponentAnswer(submission),
                        new Labs.Components.ChoiceComponentResult(correct ? 1 : 0, true),
                        function (err) {
                            if (err) {
                                // Error
                            }
                        });

                    // And finally update the UI.
                    self._showResults(correct);
                });

                // And make the view UI visible.
                $("#quiz-editor").addClass("hidden");
                $("#quiz-view").removeClass("hidden");
            });
        });
    };

    return TrueFalseQuiz;
})();
```

最后，当你连接到主机且文档准备好之后，启动测验。




```js
$(document).ready(function () {
    Labs.connect(function (err, connectionResponse) {
        if (err) {
            // ... error handling goes here ...
            return;
        }

        // Start up the true/false quiz.
        var trueFalseQuiz = new TrueFalseQuiz(connectionResponse.mode);
    });
});
```


## <a name="additional-resources"></a>其他资源
<a name="bk_addresources"> </a>


- [Office Mix 外接程序](office-mix-add-ins.md)
    
