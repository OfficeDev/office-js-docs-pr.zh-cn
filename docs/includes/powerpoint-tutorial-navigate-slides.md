本教程的这一步是，在文档幻灯片之间导航。

> [!NOTE]
> 此为 PowerPoint 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="navigate-between-slides-of-the-document"></a>在文档幻灯片之间导航

1. 在 **Home.html** 文件中，将 `TODO5` 替换为以下标记。 此标记定义在加载项任务窗格内显示的四个导航按钮。

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-first-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to First Slide</span>
        <span class="ms-Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-next-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Next Slide</span>
        <span class="ms-Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-previous-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Previous Slide</span>
        <span class="ms-Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-last-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Last Slide</span>
        <span class="ms-Button-description">Go to the last slide.</span>
    </button>
    ```

2. 在 **Home.js** 文件中，将 `TODO8` 替换为下列代码，以分配四个导航按钮的事件处理程序。

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. 在 **Home.js** 文件中，将 `TODO9` 替换为下列代码，以定义导航函数。 以下各函数均使用 `goToByIdAsync` 函数，以根据幻灯片在文档中的位置（第一张、最后一张、上一张、下一张）选择幻灯片。

    ```js
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a>测试加载项

1. 使用 Visual Studio 的同时，按 `F5` 或选择“开始”按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”加载项按钮。加载项本地托管在 IIS 上。

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. 在 PowerPoint 中，选择功能区中的“显示任务窗格”按钮，以打开加载项任务窗格。

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)


3. 使用“开始”选项卡功能区中的“新建幻灯片”按钮，将两张新幻灯片添加到文档中。 

4. 在任务窗格中，选择“前往第一张幻灯片”按钮。 此时，选择并显示文档中的第一张幻灯片。

    ![突出显示“前往第一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-first-slide.png)

5. 在任务窗格中，选择“前往下一张幻灯片”按钮。 此时，选择并显示文档中的下一张幻灯片。

    ![突出显示“前往下一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-next-slide.png)

6. 在任务窗格中，选择“前往上一张幻灯片”按钮。 此时，选择并显示文档中的上一张幻灯片。

    ![突出显示“前往上一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. 在任务窗格中，选择“前往最后一张幻灯片”按钮。 此时，选择并显示文档中的最后一张幻灯片。

    ![突出显示“前往最后一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-last-slide.png)

8. 在 Visual Studio 中，按 `Shift + F5` 或选择“停止”按钮，以停止加载项。 PowerPoint 在加载项停止时自动关闭。

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)
