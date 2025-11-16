# MainSequence Excel Add-in

This add-in lets you pull data from your backend straight into Excel
with a simple login and one custom function.

## What it does

-   Lets you sign in through a task pane.
-   Provides a function `=MainSequence.GET_DATA()` for fetching live
    data.
-   Runs inside Excel without switching apps.

## Tech Used

React, Fluent UI, Office.js, Webpack, Vercel.

## How to Run (Developers)

1.  Clone the repo.
2.  Run `npm install`.
3.  Make sure the API URL is set inside the project.
4.  Start the dev server with `npm start`.
5.  Sideload the manifest.xml in Excel from Insert \> My Add-ins \>
    Upload.

## How to Use

1.  Open the task pane and sign in.
2.  Use the function in any cell:

    =MainSequence.GET_DATA("2022-01-01", "2022-01-31", , TRUE, TRUE, 500, 0)

## Deployment

Push to GitHub to trigger a Vercel build.\
After deployment, update the production URL inside `manifest.xml`.
