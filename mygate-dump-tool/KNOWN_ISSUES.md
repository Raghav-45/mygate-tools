# Known Issues & Architectural Behaviors

## 1. AWS CloudFront Queue Timeout Limit

### Description
When requesting a multi-year dump export, MyGate's backend schedules an asynchronous background job on AWS CloudFront to generate the downloadable `.xlsx` report files.

For large residential societies with tens of thousands of helpdesk tickets per year, generating these multi-year reports on MyGate's cloud servers can occasionally take several minutes.

### Behavior & Mitigation
- The **MyGate Dump Tool** is currently configured to poll `getDownloadReportList` up to **60 times** (approximately 2 minutes at standard polling speed) waiting for MyGate's AWS CloudFront server to return a `Success` status with the S3 download link.
- If MyGate's cloud server takes longer than 2 minutes due to high server concurrency or massive ticket volume, the specific chunk will timeout and display as `"Failed / Timed out"` in the status grid.

### Recommended Workaround
If a yearly chunk times out due to extreme data volume:
1. Increase the **API Polling Speed** slider in the extension settings to `3.0s` or `4.0s` to give MyGate's cloud workers more time to process the job without hitting rate limits.
2. Alternatively, generate the export in smaller 1-year or 2-year increments during off-peak hours when MyGate's AWS queue processes reports instantaneously.
