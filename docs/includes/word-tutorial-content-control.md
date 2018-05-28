<span data-ttu-id="8297f-101">???????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-101">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span> 

> [!NOTE]
> <span data-ttu-id="8297f-p101">?????? Word ?????????????????????????????????????? [Word ?????](../tutorials/word-tutorial.yml)????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-p101">This page describes an individual step of a Word add-in tutorial. If you?ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

<span data-ttu-id="8297f-104">?????????????????? Word UI ???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-104">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties.</span></span> <span data-ttu-id="8297f-105">??????????[? Word ?????????????](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b)?</span><span class="sxs-lookup"><span data-stu-id="8297f-105">For details, see [Create forms that users complete or print in Word](https://support.office.com/en-us/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

> [!NOTE]
> <span data-ttu-id="8297f-106">????? UI ??? Word ??????????????? Word.js ????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-106">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>


## <a name="create-a-content-control"></a><span data-ttu-id="8297f-107">??????</span><span class="sxs-lookup"><span data-stu-id="8297f-107">Create a content control</span></span>

1. <span data-ttu-id="8297f-108">????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-108">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="8297f-109">???? index.html?</span><span class="sxs-lookup"><span data-stu-id="8297f-109">Open the file index.html.</span></span>
3. <span data-ttu-id="8297f-110">??? `replace-text` ??? `div` ??????????</span><span class="sxs-lookup"><span data-stu-id="8297f-110">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. <span data-ttu-id="8297f-111">?? app.js ???</span><span class="sxs-lookup"><span data-stu-id="8297f-111">Open the app.js file.</span></span>

5. <span data-ttu-id="8297f-112">?? `insert-table` ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-112">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="8297f-113">? `insertTable` ????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-113">Below the `insertTable` function, add the following function:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="8297f-p103">? `TODO1` ???????????</span><span class="sxs-lookup"><span data-stu-id="8297f-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="8297f-116">????????????????Office 365??</span><span class="sxs-lookup"><span data-stu-id="8297f-116">This code is intended to wrap the phrase "Office 365" in a content control.</span></span> <span data-ttu-id="8297f-117">?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-117">It makes a simplifying assumption that the string is present and the user has selected it.</span></span>
   - <span data-ttu-id="8297f-118">??????????????`ContentControl.title`</span><span class="sxs-lookup"><span data-stu-id="8297f-118">The `ContentControl.title` property specifies the visible title of the content control.</span></span> 
   - <span data-ttu-id="8297f-119">???????????? `ContentControlCollection.getByTag` ????????????????????????`ContentControl.tag`</span><span class="sxs-lookup"><span data-stu-id="8297f-119">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span> 
   - <span data-ttu-id="8297f-120">??????????`ContentControl.appearance`</span><span class="sxs-lookup"><span data-stu-id="8297f-120">The `ContentControl.appearance` property specifies the visual look of the control.</span></span> <span data-ttu-id="8297f-121">????Tags??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-121">Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title.</span></span> <span data-ttu-id="8297f-122">????????BoundingBox???None??</span><span class="sxs-lookup"><span data-stu-id="8297f-122">Other possible values are "BoundingBox" and "None".</span></span>
   - <span data-ttu-id="8297f-123">????????????????`ContentControl.color`</span><span class="sxs-lookup"><span data-stu-id="8297f-123">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="8297f-124">?????????</span><span class="sxs-lookup"><span data-stu-id="8297f-124">Replace the content of the content control</span></span>

1. <span data-ttu-id="8297f-125">???? index.html?</span><span class="sxs-lookup"><span data-stu-id="8297f-125">Open the file index.html.</span></span>
3. <span data-ttu-id="8297f-126">??? `create-content-control` ??? `div` ??????????</span><span class="sxs-lookup"><span data-stu-id="8297f-126">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>
    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

4. <span data-ttu-id="8297f-127">?? app.js ???</span><span class="sxs-lookup"><span data-stu-id="8297f-127">Open the app.js file.</span></span>

5. <span data-ttu-id="8297f-128">?? `create-content-control` ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-128">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

6. <span data-ttu-id="8297f-129">? `createContentControl` ????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-129">Below the `createContentControl` function, add the following function:</span></span>

    <span data-ttu-id="8297f-130">\`\`\`js    function replaceContentInControl() {      Word.run(function (context) {</span><span class="sxs-lookup"><span data-stu-id="8297f-130">\`\`\`js    function replaceContentInControl() {      Word.run(function (context) {</span></span>
            
            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    <span data-ttu-id="8297f-131">}</span><span class="sxs-lookup"><span data-stu-id="8297f-131"></span></span>
    ``` 

7. Replace `TODO1` with the following code. 
    > [!NOTE]
    > The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag. We use `getFirst` to get a reference to the desired control.

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="8297f-132">?????</span><span class="sxs-lookup"><span data-stu-id="8297f-132">Test the add-in</span></span>

1. <span data-ttu-id="8297f-133">?????????? Git Bash ?????? Node.JS ?????????????????? Ctrl+C ?????????? Web ????</span><span class="sxs-lookup"><span data-stu-id="8297f-133">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="8297f-134">????? Git Bash ?????? Node.JS ???????????????????****????</span><span class="sxs-lookup"><span data-stu-id="8297f-134">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
     > [!NOTE]
     > <span data-ttu-id="8297f-135">????????????? app.js ???????????????????????????????????? JavaScript?????????????????? app.js ??????????</span><span class="sxs-lookup"><span data-stu-id="8297f-135">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="8297f-136">?????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-136">In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command.</span></span> <span data-ttu-id="8297f-137">??????????</span><span class="sxs-lookup"><span data-stu-id="8297f-137">After the build, restart the server.</span></span> <span data-ttu-id="8297f-138">?????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-138">The next few steps carry out this process.</span></span>
2. <span data-ttu-id="8297f-139">???? `npm run build`??? ES6 ??????????? Office ??????????? JavaScript?</span><span class="sxs-lookup"><span data-stu-id="8297f-139">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>
3. <span data-ttu-id="8297f-140">???? `npm start`????? localhost ???? Web ????</span><span class="sxs-lookup"><span data-stu-id="8297f-140">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="8297f-141">??????????????????????****????????????****??????????</span><span class="sxs-lookup"><span data-stu-id="8297f-141">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="8297f-142">???????????????****????????????Office 365?????</span><span class="sxs-lookup"><span data-stu-id="8297f-142">In the taskpane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>
6. <span data-ttu-id="8297f-143">??????????????Office 365?????????????****???</span><span class="sxs-lookup"><span data-stu-id="8297f-143">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button.</span></span> <span data-ttu-id="8297f-144">????????????????????????</span><span class="sxs-lookup"><span data-stu-id="8297f-144">Note that the phrase is wrapped in tags labelled "Service Name".</span></span>
7. <span data-ttu-id="8297f-145">?????????****??????????????????Fabrikam Online Productivity Suite??</span><span class="sxs-lookup"><span data-stu-id="8297f-145">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Word ?? - ????????????](../images/word-tutorial-content-control.png)
