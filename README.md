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

### GCP Bucket config
Here is the summary of the commands to run in your Google Cloud Shell (or local terminal) to permanently fix CORS and file type issues for your Excel Add-in bucket.

Replace YOUR_BUCKET_NAME with your actual bucket (e.g., mainsequence-excel-addin-development).

#### Step 1: Create the CORS configuration file
Run this single command to create a file named cors.json with permissive settings (allows all origins) in your current folder.

```Bash

echo '[{"origin": ["*"],"method": ["GET", "HEAD", "OPTIONS"],"responseHeader": ["*"],"maxAgeSeconds": 3600}]' > cors.json
```
#### Step 2: Apply the configuration to the bucket
This pushes the rule to Google Cloud.

```Bash

gcloud storage buckets update gs://YOUR_BUCKET_NAME --cors-file=cors.json
```
#### Step 3: Fix Content Types (The "Hidden" Fix)
Even with CORS fixed, Excel will reject your files if Google thinks they are "text/plain". Run these to force them to be executable code.

```Bash

# Fix the JSON manifest/metadata
gsutil setmeta -h "Content-Type:application/json" gs://YOUR_BUCKET_NAME/web/functions.json

# Fix the JavaScript logic
gsutil setmeta -h "Content-Type:application/javascript" gs://YOUR_BUCKET_NAME/web/functions.js
```
Step 4: Verify it worked
Run this "fake request" to see if the bucket responds with the correct permission headers.

```Bash

curl -I -H "Origin: https://excel.officeapps.live.com" https://storage.googleapis.com/YOUR_BUCKET_NAME/web/functions.json
```
Success Indicator: You should see this line in the output: access-control-allow-origin: *