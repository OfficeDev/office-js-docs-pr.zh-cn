# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="cd952-101">?? Angular ?? Excel ???</span><span class="sxs-lookup"><span data-stu-id="cd952-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="cd952-102">??????????? Angular ? Excel JavaScript API ?? Excel ???????</span><span class="sxs-lookup"><span data-stu-id="cd952-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cd952-103">????</span><span class="sxs-lookup"><span data-stu-id="cd952-103">Prerequisites</span></span>

- <span data-ttu-id="cd952-104">?????? [Angular CLI ????](https://github.com/angular/angular-cli#prerequisites)??????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="cd952-105">???? [Angular CLI](https://github.com/angular/angular-cli)?</span><span class="sxs-lookup"><span data-stu-id="cd952-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="cd952-106">??????? [Yeoman](https://github.com/yeoman/yo) ? [Office ???? Yeoman ???](https://github.com/OfficeDev/generator-office)?</span><span class="sxs-lookup"><span data-stu-id="cd952-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="cd952-107">???? Angular ??</span><span class="sxs-lookup"><span data-stu-id="cd952-107">Generate a new Angular app</span></span>

<span data-ttu-id="cd952-108">?? Angular CLI ?? Angular ???</span><span class="sxs-lookup"><span data-stu-id="cd952-108">Use the Angular CLI to generate your Angular app.</span></span> <span data-ttu-id="cd952-109">??????????</span><span class="sxs-lookup"><span data-stu-id="cd952-109">From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="cd952-110">??????</span><span class="sxs-lookup"><span data-stu-id="cd952-110">Generate the manifest file</span></span>

<span data-ttu-id="cd952-111">???????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="cd952-112">??????????</span><span class="sxs-lookup"><span data-stu-id="cd952-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="cd952-113">?? Yeoman ?????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-113">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="cd952-114">?????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-114">Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office
    ```
    - <span data-ttu-id="cd952-115">**??????????????** `No`</span><span class="sxs-lookup"><span data-stu-id="cd952-115">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="cd952-116">**??????????????:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="cd952-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="cd952-117">**?????? Office ????????:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="cd952-117">**Which Office client application would you like to support?:** `Excel`</span></span>
    - <span data-ttu-id="cd952-118">**??????????** `No`</span><span class="sxs-lookup"><span data-stu-id="cd952-118">**Would you like to create a new add-in?:** `No`</span></span>

    <span data-ttu-id="cd952-p103">???????????????resource.html?****???????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-p103">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Yeoman ???](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="cd952-123">???????? **package.json**???????****??????</span><span class="sxs-lookup"><span data-stu-id="cd952-123">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="cd952-124">??????</span><span class="sxs-lookup"><span data-stu-id="cd952-124">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="cd952-125">????????????**??? Office ????? Yeoman ???**??????</span><span class="sxs-lookup"><span data-stu-id="cd952-125">For this quickstart, you can use the certificates that the **Yeoman generator for Office Add-ins** provides.</span></span> <span data-ttu-id="cd952-126">????????????????????**????**???????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-126">You've already installed the generator globally (as part of the **Prerequisites** for this quickstart), so you'll just need to copy the certificates from the global install location into your app folder.</span></span> <span data-ttu-id="cd952-127">???????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-127">The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="cd952-128">???????????????????? **npm** ??????</span><span class="sxs-lookup"><span data-stu-id="cd952-128">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="cd952-129">????????????????????? **npm** ??????</span><span class="sxs-lookup"><span data-stu-id="cd952-129">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="cd952-130">??????????? `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` ????</span><span class="sxs-lookup"><span data-stu-id="cd952-130">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder.</span></span> <span data-ttu-id="cd952-131">????? `certs` ??????????</span><span class="sxs-lookup"><span data-stu-id="cd952-131">From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="cd952-132">?????????? 1 ???? Angular ???????????? `certs` ????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-132">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="cd952-133">??????</span><span class="sxs-lookup"><span data-stu-id="cd952-133">Update the app</span></span>

1. <span data-ttu-id="cd952-134">????????????????? **package.json**?</span><span class="sxs-lookup"><span data-stu-id="cd952-134">In your code editor, open **package.json** in the root of the project.</span></span> <span data-ttu-id="cd952-135">? `start` ????????????? SSL ??? 3000 ?????????</span><span class="sxs-lookup"><span data-stu-id="cd952-135">Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="cd952-136">????????? **.angular-cli.json**?</span><span class="sxs-lookup"><span data-stu-id="cd952-136">Open **.angular-cli.json** in the root of the project.</span></span> <span data-ttu-id="cd952-137">? **defaults** ????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-137">Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

    ```json
    "defaults": {
      "styleExt": "css",
      "component": {},
      "serve": {
        "sslKey": "certs/server.key",
        "sslCert": "certs/server.crt"
      }
    }
    ```

3. <span data-ttu-id="cd952-138">?? **src/index.html**??? `</head>` ???????? `<script>` ??????????</span><span class="sxs-lookup"><span data-stu-id="cd952-138">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="cd952-139">???src/main.ts?****?? `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` ???????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-139">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="cd952-140">???src/polyfills.ts?****???????? `import` ???????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-140">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="cd952-141">??src/polyfills.ts?****???????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-141">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

7. <span data-ttu-id="cd952-142">???src/app/app.component.html?****??????????? HTML????????</span><span class="sxs-lookup"><span data-stu-id="cd952-142">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onSetColor()">Set color</button>
        </div>
    </div>
    ```

8. <span data-ttu-id="cd952-143">???src/app/app.component.css?****??????????? CSS ??????????</span><span class="sxs-lookup"><span data-stu-id="cd952-143">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

9. <span data-ttu-id="cd952-144">?? **src/app/app.component.ts**?????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-144">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onSetColor() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="cd952-145">?????????</span><span class="sxs-lookup"><span data-stu-id="cd952-145">Start the dev server</span></span>

1. <span data-ttu-id="cd952-146">???????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-146">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="cd952-p108">? Web ??????? `https://localhost:3000`???????????????????????????????????????????[????????????????](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)?</span><span class="sxs-lookup"><span data-stu-id="cd952-p108">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="cd952-150">Chrome?Web ?????????????????????????[????????????????](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-150">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span> <span data-ttu-id="cd952-151">???? Chrome ????????? Internet Explorer ? Microsoft Edge ?? `https://localhost:3000`????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-151">You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="cd952-152">?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-152">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="cd952-153">??</span><span class="sxs-lookup"><span data-stu-id="cd952-153">Try it out</span></span>

1. <span data-ttu-id="cd952-154">?????????? Excel ????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-154">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="cd952-155">Windows?[? Windows ???? Office ???](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="cd952-155">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="cd952-156">Excel Online?[? Office Online ???? Office ???](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="cd952-156">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="cd952-157">iPad ? Mac?[? iPad ? Mac ???? Office ???](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="cd952-157">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="cd952-158">? Excel ??????????****?????????????????****??????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-158">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel ?????](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="cd952-160">????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-160">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="cd952-161">???????????????****?????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-161">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel ???](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="cd952-163">????</span><span class="sxs-lookup"><span data-stu-id="cd952-163">Next steps</span></span>

<span data-ttu-id="cd952-p110">?????? Angular ???? Excel ????????????? Excel ????????? Excel ????????????????????</span><span class="sxs-lookup"><span data-stu-id="cd952-p110">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="cd952-166">Excel ?????</span><span class="sxs-lookup"><span data-stu-id="cd952-166">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="cd952-167">????</span><span class="sxs-lookup"><span data-stu-id="cd952-167">See also</span></span>

* [<span data-ttu-id="cd952-168">Excel ?????</span><span class="sxs-lookup"><span data-stu-id="cd952-168">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="cd952-169">Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="cd952-169">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="cd952-170">Excel ???????</span><span class="sxs-lookup"><span data-stu-id="cd952-170">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="cd952-171">Excel JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="cd952-171">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
