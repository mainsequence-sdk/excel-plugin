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
   start_date = cell reference
   end_date = cell reference
   unique_identifier_list = empty cell reference for now getting data from `get_data_between_dates_from_node_identifier` API
   update_hash = can bee a empty cell reference or any string to complete the parameters it's an optional currently hardcoded in code as `null`

   `=MainSequence.GET_DATA(C8, D8, A10:A11,E8, TRUE, TRUE, 500, 0,"null")`

## Deployment

Push to GitHub to trigger a Vercel build.\
After deployment, update the production URL inside `manifest.xml`.
