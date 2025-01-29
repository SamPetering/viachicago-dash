### quickstart

1. clone repo: `git clone https://github.com/SamPetering/viachicago-dash.git && cd viachicago-dash`
2. install dependencies: `pnpm i`
3. create `.clasp.json` (modify `.clasp.example.json`)
   - replace `scriptId` with the apps script id for the project (found in Your Google Sheet > Extensions > Apps Script > Project Settings > IDs > Script ID)
   - replace `rootDir` with the path to the root of this project on your machine.
4. authorize clasp with your google account: `clasp login`
5. build and push to apps script project `pnpm bp`
