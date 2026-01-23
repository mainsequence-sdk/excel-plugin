# MainSequence Excel Add-in

This add-in lets you pull data from your backend straight into Excel
with a simple login and one custom function.

## What it does

-   Lets you sign in through a task pane.
-   Provides a function `=MainSequence.GET_DATA()` for fetching live
    data.
-   Runs inside Excel without switching apps.

## Tech Used

React, Fluent UI, Office.js, Webpack.

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

Production builds rewrite `manifest.xml` with the GCS base URL and upload the built assets to the bucket.
GCS base: `https://storage.googleapis.com/tsorm-production/excel-addin/`
Manifest URL: `https://storage.googleapis.com/tsorm-production/excel-addin/manifest.xml`
If you want a friendly URL (e.g. `https://main-sequence.app/excel-addin`), have Django redirect to that manifest.

### GCP Bucket config
Here is the summary of the commands to run in your Google Cloud Shell (or local terminal) to permanently fix CORS and file type issues for your Excel Add-in bucket.

Bucket: `gs://tsorm-production` (assets live under `/excel-addin/`).

#### Step 1: Create the CORS configuration file
Run this single command to create a file named cors.json with permissive settings (allows all origins) in your current folder.

```Bash

echo '[{"origin": ["*"],"method": ["GET", "HEAD", "OPTIONS"],"responseHeader": ["*"],"maxAgeSeconds": 3600}]' > cors.json
```
#### Step 2: Apply the configuration to the bucket
This pushes the rule to Google Cloud.

```Bash

gcloud storage buckets update gs://tsorm-production --cors-file=cors.json
```
#### Step 3: Fix Content Types (The "Hidden" Fix)
Even with CORS fixed, Excel will reject your files if Google thinks they are "text/plain". Run these to force them to be executable code.

```Bash

# Fix the JSON manifest/metadata
gsutil setmeta -h "Content-Type:application/json" gs://tsorm-production/excel-addin/functions.json

# Fix the JavaScript logic
gsutil setmeta -h "Content-Type:application/javascript" gs://tsorm-production/excel-addin/functions.js
```
Step 4: Verify it worked
Run this "fake request" to see if the bucket responds with the correct permission headers.

```Bash

curl -I -H "Origin: https://excel.officeapps.live.com" https://storage.googleapis.com/tsorm-production/excel-addin/functions.json
```
Success Indicator: You should see this line in the output: access-control-allow-origin: *
