# Known Issues & Architectural Behaviors

## 1. Cloud Export Completion Timeout

### Description
When requesting a multi-year dump export, MyGate's backend schedules asynchronous background jobs to generate the downloadable `.xlsx` report files, one date range at a time.

For large residential societies with tens of thousands of helpdesk tickets per year, generating these multi-year reports on MyGate's cloud servers can occasionally take several minutes.

### Behavior & Mitigation
- The **MyGate Dump Tool** polls `getDownloadReportList` up to **45 times**, with a default **2-second delay** between checks (about 90 seconds of waiting, plus request time). It accepts only a new successful entry with matching dates and statuses, and rejects reused cloud-file links.
- If the new export does not become ready in time, the run displays `Stopped`. No partial Master Dump is saved, no later date ranges are requested, and dependent Analysis is blocked. Files already downloaded by other selected tools remain available.
- Each date range gets one export request. Status checks do not generate reports; there are no automatic export retries.

### Recommended Workaround
If a yearly chunk times out, inspect the existing request on MyGate's Generated Reports page and allow it to finish rather than immediately requesting another export. The stopped run does not resume automatically.

For a future run of the standalone Dump Tool, its **API Polling Speed** setting can lengthen the interval between status checks. Handoff keeps the original 2-second default.
