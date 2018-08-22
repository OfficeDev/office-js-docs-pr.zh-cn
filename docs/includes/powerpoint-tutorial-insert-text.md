本教程的这一步是，向包含一天中[必应](https://www.bing.com)照片的标题幻灯片添加文本。

> [!NOTE]
> 此为 PowerPoint 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="add-text-to-a-slide"></a>向幻灯片添加文本 

1. 在 **Home.html** 文件中，将 `TODO3` 替换为以下标记。 此标记定义在加载项任务窗格内显示的“插入文本”**** 按钮。

    ```html
        <br /><br />
        <button class="ms-Button ms-Button--primary" id="insert-text">
            <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="ms-Button-label">Insert Text</span>
            <span class="ms-Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. 在 **Home.js** 文件中，将 `TODO4` 替换为下列代码，以分配“插入文本”**** 按钮的事件处理程序。

    ```js
    $('#insert-text').click(insertText);
    ```

3. 在 **Home.js** 文件中，将 `TODO5` 替换为下列代码，以定义 **insertText** 函数。 此函数将文本插入当前幻灯片。

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a>测试加载项

1. 使用 Visual Studio 的同时，按 `F5` 或选择“开始”**** 按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. 在 PowerPoint 中，选择功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. 在任务窗格中，选择“插入图像”**** 按钮，将一天中的必应照片添加到当前幻灯片，再为包含标题文本框的幻灯片选择一种设计。

    ![突出显示“插入图像”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. 将光标置于标题幻灯片上的文本框中，再选择任务窗格中的“插入文本”**** 按钮，向幻灯片添加文本。

    ![突出显示“插入文本”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-text.png)


5. 在 Visual Studio 中，按 `Shift + F5` 或选择“停止”**** 按钮，以停止加载项。 PowerPoint 在加载项停止时自动关闭。

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)