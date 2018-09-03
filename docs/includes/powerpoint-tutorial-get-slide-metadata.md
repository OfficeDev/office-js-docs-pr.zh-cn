<span data-ttu-id="e1329-101">本教程的这一步是，检索选定幻灯片的元数据。</span><span class="sxs-lookup"><span data-stu-id="e1329-101">In this step of the tutorial, you'll retrieve metadata for the selected slide.</span></span>

> [!NOTE]
> <span data-ttu-id="e1329-102">此为 PowerPoint 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="e1329-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="e1329-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="e1329-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="get-slide-metadata"></a><span data-ttu-id="e1329-104">获取幻灯片元数据</span><span class="sxs-lookup"><span data-stu-id="e1329-104">Get slide metadata</span></span>

1. <span data-ttu-id="e1329-105">在 **Home.html** 文件中，将 `TODO4` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="e1329-105">In the **Home.html** file, replace `TODO4` with the following markup.</span></span> <span data-ttu-id="e1329-106">此标记定义在加载项任务窗格内显示的“获取幻灯片元数据”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="e1329-106">This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="get-slide-metadata">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Get Slide Metadata</span>
        <span class="ms-Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="e1329-107">在 **Home.js** 文件中，将 `TODO6` 替换为下列代码，以分配“获取幻灯片元数据”**** 按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e1329-107">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="e1329-108">在 **Home.js** 文件中，将 `TODO7` 替换为下列代码，以定义 **getSlideMetadata** 函数。</span><span class="sxs-lookup"><span data-stu-id="e1329-108">In the **Home.js** file, replace `TODO7` with the following code to define the **getSlideMetadata** function.</span></span> <span data-ttu-id="e1329-109">此函数检索选定一张或多张幻灯片的元数据，并将它写入加载项任务窗格内的弹出对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="e1329-109">This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="e1329-110">测试加载项</span><span class="sxs-lookup"><span data-stu-id="e1329-110">Test the add-in</span></span>

1. <span data-ttu-id="e1329-p104">使用 Visual Studio 的同时，按 `F5` 或选择“开始”**** 按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="e1329-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="e1329-114">在 PowerPoint 中，选择功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="e1329-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="e1329-116">在任务窗格中，选择“获取幻灯片元数据”**** 按钮，以获取选定幻灯片的元数据。</span><span class="sxs-lookup"><span data-stu-id="e1329-116">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide.</span></span> <span data-ttu-id="e1329-117">此时，幻灯片元数据写入到任务窗格底部的弹出对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="e1329-117">The slide metadata is written to the popup dialog window at the bottom of the task pane.</span></span> <span data-ttu-id="e1329-118">在此示例中，JSON 元数据中的 `slides` 数组包含一个对象，用于指定选定幻灯片的 `id`、`title` 和 `index`。</span><span class="sxs-lookup"><span data-stu-id="e1329-118">In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide.</span></span> <span data-ttu-id="e1329-119">如果在检索幻灯片元数据时选择了多张幻灯片，JSON 元数据中的 `slides` 数组会对每张选定幻灯片都包含一个对象。</span><span class="sxs-lookup"><span data-stu-id="e1329-119">If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![突出显示“获取幻灯片元数据”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="e1329-121">在 Visual Studio 中，按 `Shift + F5` 或选择“停止”**** 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="e1329-121">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="e1329-122">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="e1329-122">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)
