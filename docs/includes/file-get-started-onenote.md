# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="4a4bb-101">???? OneNote ???</span><span class="sxs-lookup"><span data-stu-id="4a4bb-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="4a4bb-102">??????????? jQuery ? Office JavaScript API ?? OneNote ????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4a4bb-103">????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-103">Prerequisites</span></span>

- [<span data-ttu-id="4a4bb-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="4a4bb-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="4a4bb-105">??????? [Yeoman](https://github.com/yeoman/yo) ? [Office ???? Yeoman ???](https://github.com/OfficeDev/generator-office)?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="4a4bb-106">???????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-106">Create the add-in project</span></span>

1. <span data-ttu-id="4a4bb-107">????????????????????`my-onenote-addin`??</span><span class="sxs-lookup"><span data-stu-id="4a4bb-107">Create a folder on your local drive and name it `my-onenote-addin`.</span></span> <span data-ttu-id="4a4bb-108">?????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="4a4bb-109">???????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-onenote-addin
    ```

3. <span data-ttu-id="4a4bb-110">?? Yeoman ????? OneNote ??????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-110">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="4a4bb-111">?????????????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="4a4bb-112">**???????????????:** `No`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-112">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="4a4bb-113">**??????????????:** `OneNote Add-in`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-113">**What do you want to name your add-in?:** `OneNote Add-in`</span></span>
    - <span data-ttu-id="4a4bb-114">**?????? Office ????????:** `OneNote`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-114">**Which Office client application would you like to support?:** `OneNote`</span></span>
    - <span data-ttu-id="4a4bb-115">**??????????:** `Yes`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-115">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="4a4bb-116">**????? TypeScript?:** `No`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-116">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="4a4bb-117">**?????** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-117">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="4a4bb-p103">???????????????resource.html?****???????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![?? Yeoman ?????????????](../images/yo-office-onenote-jquery.png)


## <a name="update-the-code"></a><span data-ttu-id="4a4bb-122">????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-122">Update the code</span></span>

1. <span data-ttu-id="4a4bb-123">??????????????????index.html?****?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-123">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="4a4bb-124">????????????????? HTML?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-124">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="4a4bb-125">? `<body>` ???? `<main>` ????????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-125">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span> <span data-ttu-id="4a4bb-126">???? [Office UI Fabric ??](http://dev.office.com/fabric/components)??????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-126">This adds a text area and a button using [Office UI Fabric components](http://dev.office.com/fabric/components).</span></span>

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. <span data-ttu-id="4a4bb-127">?????app.js?****??????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-127">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="4a4bb-128">???????????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-128">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="4a4bb-129">????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-129">Update the manifest</span></span>

1. <span data-ttu-id="4a4bb-130">?????one-note-add-in-manifest.xml?****??????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-130">Open the file **one-note-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="4a4bb-131">?????????`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-131">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="4a4bb-132">??????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-132">Replace it with your name.</span></span>

3. <span data-ttu-id="4a4bb-133">??? `DefaultValue` ???????`Description`</span><span class="sxs-lookup"><span data-stu-id="4a4bb-133">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="4a4bb-134">??????A task pane add-in for OneNote?****?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-134">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="4a4bb-135">?????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-135">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="4a4bb-136">?????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-136">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="4a4bb-137">??</span><span class="sxs-lookup"><span data-stu-id="4a4bb-137">Try it out</span></span>

1. <span data-ttu-id="4a4bb-138">? [OneNote Online](https://www.onenote.com/notebooks) ??????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-138">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="4a4bb-139">????????>?Office ????****????Office ????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-139">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="4a4bb-140">????????????????????????****?????????????****?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-140">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="4a4bb-141">?????????????????????????****?????????????****?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-141">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="4a4bb-142">???????????????????****????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-142">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="4a4bb-143">??????????????????????**?one-note-add-in-manifest.xml?**????**????**?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-143">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="4a4bb-144">??**??**?????????????**??????**????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-144">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="4a4bb-145">?????? OneNote ???? iFrame ????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-145">6- The add-in opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="4a4bb-146">??????????????????**????**??</span><span class="sxs-lookup"><span data-stu-id="4a4bb-146">Enter some text in the text area and then choose **Add outline**.</span></span> <span data-ttu-id="4a4bb-147">?????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-147">The text you entered is added to the page.</span></span> 

    ![???????? OneNote ???](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="4a4bb-149">???????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-149">Troubleshooting and tips</span></span>

- <span data-ttu-id="4a4bb-p111">???????????????????????? Internet Explorer ? Chrome ??? Gulp Web ???????????????????????????????? iFrame?</span><span class="sxs-lookup"><span data-stu-id="4a4bb-p111">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="4a4bb-p112">?? OneNote ???????????????????????????????**??? `_proto_` ?????????????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-p112">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![????????? OneNote ??](../images/onenote-debug.png)

- <span data-ttu-id="4a4bb-p113">???????????? HTTP ??????????????????????????????? HTTPS ???</span><span class="sxs-lookup"><span data-stu-id="4a4bb-p113">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="4a4bb-158">??????????????????????????????????????IFrame ?????????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-158">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="4a4bb-159">????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-159">Next steps</span></span>

<span data-ttu-id="4a4bb-160">???????? OneNote ????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-160">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="4a4bb-161">???????????? OneNote ???????????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-161">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="4a4bb-162">OneNote JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-162">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="4a4bb-163">????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-163">See also</span></span>

- [<span data-ttu-id="4a4bb-164">OneNote JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-164">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="4a4bb-165">OneNote JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="4a4bb-165">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="4a4bb-166">Rubric Grader ??</span><span class="sxs-lookup"><span data-stu-id="4a4bb-166">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="4a4bb-167">Office ???????</span><span class="sxs-lookup"><span data-stu-id="4a4bb-167">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
