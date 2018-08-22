<span data-ttu-id="a453f-101">本教程的这一步是，在文档幻灯片之间导航。</span><span class="sxs-lookup"><span data-stu-id="a453f-101">In this step of the tutorial, you'll navigate between the slides of a document.</span></span>

> [!NOTE]
> <span data-ttu-id="a453f-102">此为 PowerPoint 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="a453f-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="a453f-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="a453f-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="navigate-between-slides-of-the-document"></a><span data-ttu-id="a453f-104">在文档幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="a453f-104">Navigate between slides of the document</span></span>

1. <span data-ttu-id="a453f-105">在 **Home.html** 文件中，将 `TODO5` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="a453f-105">In the **Home.html** file, replace `TODO5` with the following markup.</span></span> <span data-ttu-id="a453f-106">此标记定义在加载项任务窗格内显示的四个导航按钮。</span><span class="sxs-lookup"><span data-stu-id="a453f-106">This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

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

2. <span data-ttu-id="a453f-107">在 **Home.js** 文件中，将 `TODO8` 替换为下列代码，以分配四个导航按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="a453f-107">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="a453f-108">在 **Home.js** 文件中，将 `TODO9` 替换为下列代码，以定义导航函数。</span><span class="sxs-lookup"><span data-stu-id="a453f-108">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="a453f-109">以下各函数均使用 `goToByIdAsync` 函数，以根据幻灯片在文档中的位置（第一张、最后一张、上一张、下一张）选择幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a453f-109">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, next).</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="a453f-110">测试加载项</span><span class="sxs-lookup"><span data-stu-id="a453f-110">Test the add-in</span></span>

1. <span data-ttu-id="a453f-p104">使用 Visual Studio 的同时，按 `F5` 或选择“开始”**** 按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="a453f-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="a453f-114">在 PowerPoint 中，选择功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="a453f-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)


3. <span data-ttu-id="a453f-116">使用“开始”**** 选项卡功能区中的“新建幻灯片”**** 按钮，将两张新幻灯片添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="a453f-116">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span> 

4. <span data-ttu-id="a453f-117">在任务窗格中，选择“前往第一张幻灯片”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="a453f-117">In the task pane, choose the **Go to First Slide** button.</span></span> <span data-ttu-id="a453f-118">此时，选择并显示文档中的第一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a453f-118">The first slide in the document is selected and displayed.</span></span>

    ![突出显示“前往第一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="a453f-120">在任务窗格中，选择“前往下一张幻灯片”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="a453f-120">In the task pane, choose the **Go to Next Slide** button.</span></span> <span data-ttu-id="a453f-121">此时，选择并显示文档中的下一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a453f-121">The next slide in the document is selected and displayed.</span></span>

    ![突出显示“前往下一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="a453f-123">在任务窗格中，选择“前往上一张幻灯片”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="a453f-123">In the task pane, choose the **Go to Previous Slide** button.</span></span> <span data-ttu-id="a453f-124">此时，选择并显示文档中的上一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a453f-124">The previous slide in the document is selected and displayed.</span></span>

    ![突出显示“前往上一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="a453f-126">在任务窗格中，选择“前往最后一张幻灯片”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="a453f-126">In the task pane, choose the **Go to Last Slide** button.</span></span> <span data-ttu-id="a453f-127">此时，选择并显示文档中的最后一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a453f-127">The last slide in the document is selected and displayed.</span></span>

    ![突出显示“前往最后一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="a453f-129">在 Visual Studio 中，按 `Shift + F5` 或选择“停止”**** 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="a453f-129">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="a453f-130">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="a453f-130">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)
