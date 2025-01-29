### quickstart

clone repo: `git clone https://github.com/SamPetering/viachicago-dash.git && cd viachicago-dash`

install dependencies: `pnpm i`

inside `.clasp.json`, replace `scriptId` with the project's apps script id
- (found in Your Google Sheet > Extensions > Apps Script > Project Settings > IDs > Script ID)

authorize clasp with your google account: `clasp login`

build and push `pnpm bp`
