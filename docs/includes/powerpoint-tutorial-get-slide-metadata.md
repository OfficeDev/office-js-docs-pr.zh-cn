本教程的这一步是，检索选定幻灯片的元数据。

> [!NOTE]
> 此为 PowerPoint 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="get-slide-metadata"></a>获取幻灯片元数据

1. 在 **Home.html** 文件中，将 `TODO4` 替换为以下标记。 此标记定义在加载项任务窗格内显示的“获取幻灯片元数据”**** 按钮。

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="get-slide-metadata">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Get Slide Metadata</span>
        <span class="ms-Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. 在 **Home.js** 文件中，将 `TODO6` 替换为下列代码，以分配“获取幻灯片元数据”**** 按钮的事件处理程序。

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. 在 **Home.js** 文件中，将 `TODO7` 替换为下列代码，以定义 **getSlideMetadata** 函数。 此函数检索选定一张或多张幻灯片的元数据，并将它写入加载项任务窗格内的弹出对话框窗口。

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

## <a name="test-the-add-in"></a>测试加载项

1. 使用 Visual Studio 的同时，按 `F5` 或选择“开始”**** 按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. 在 PowerPoint 中，选择功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. 在任务窗格中，选择“获取幻灯片元数据”**** 按钮，以获取选定幻灯片的元数据。 此时，幻灯片元数据写入到任务窗格底部的弹出对话框窗口。 在此示例中，JSON 元数据中的 `slides` 数组包含一个对象，用于指定选定幻灯片的 `id`、`title` 和 `index`。 如果在检索幻灯片元数据时选择了多张幻灯片，JSON 元数据中的 `slides` 数组会对每张选定幻灯片都包含一个对象。

    ![突出显示“获取幻灯片元数据”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-get-slide-metadata.png)

4. 在 Visual Studio 中，按 `Shift + F5` 或选择“停止”**** 按钮，以停止加载项。 PowerPoint 在加载项停止时自动关闭。

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)
