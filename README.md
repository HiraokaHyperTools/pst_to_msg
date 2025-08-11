# @hiraokahypertools/pst_to_msg

[![npm](https://img.shields.io/npm/v/@hiraokahypertools/pst_to_msg)](https://www.npmjs.com/package/@hiraokahypertools/pst_to_msg)

Convert messages in Outlook PST files into MSG format

Links: [API Documentation](https://hiraokahypertools.github.io/pst_to_msg/typedoc/)

## Installation

```bash
npm install @hiraokahypertools/pst_to_msg
```

## Usage

### Basic Usage

```js
const { convertToMsg } = require('@hiraokahypertools/pst_to_msg');
const { openPstFile } = require('@hiraokahypertools/pst-extractor');
const { writeFileSync } = require('fs');

async function convertSingleMessage() {
  const pstFile = await openPstFile('example.pst', { ansiEncoding: "cp932" });

  try {
    const unnamedFolder = await pstFile.getRootFolder();
    const topFolder = await unnamedFolder.getSubFolder(0);
    const message = await topFolder.getEmail(0); // Get first message

    const msgBuffer = await convertToMsg(message, pstFile);
    writeFileSync('output.msg', msgBuffer);

    console.log('Message converted successfully!');
  } finally {
    pstFile.close();
  }
}

convertSingleMessage();

```

### Batch Export Messages

Traverse the entire PST file and extract all found messages into file system folders:

```js
const { convertToMsg } = require('@hiraokahypertools/pst_to_msg');
const { openPstFile } = require('@hiraokahypertools/pst-extractor');
const { writeFile, mkdir } = require('node:fs/promises');

async function runBatchConversion() {
  const extractPstFileTo = async (pstFilePath, outDirRoot) => {
    const pstFile = await openPstFile(pstFilePath, { ansiEncoding: "cp932" });

    try {
      async function walk(folder, outDir) {
        // Process subfolders
        const subFolderCount = await folder.getSubFolderCount();
        for (let i = 0; i < subFolderCount; i++) {
          const subFolder = await folder.getSubFolder(i);
          await walk(subFolder, `${outDir}${normFileName(subFolder.displayName)}/`);
        }

        // Process emails in current folder
        const emailCount = await folder.getEmailCount();
        for (let i = 0; i < emailCount; i++) {
          const message = await folder.getEmail(i);
          const msg = await convertToMsg(message, pstFile);

          await mkdir(outDir, { recursive: true });
          const saveTo = `${outDir}${normFileName(message.subject || 'untitled')}.msg`;
          await writeFile(saveTo, msg);
        }
      }

      const rootFolder = (await pstFile.getRootFolder());
      await walk(rootFolder, outDirRoot ? outDirRoot + "/" : "");
    } finally {
      pstFile.close();
    }
  };

  // Utility function to normalize file names
  function normFileName(name) {
    return name.replace(/[<>:"/\\|?*]/g, '_').substring(0, 100);
  }

  // Usage
  await extractPstFileTo('input.pst', 'output/messages');
}

runBatchConversion();
```

## Development

```bash
# Install dependencies
npm install

# Build the project
npm run build

# Run tests
npm test

# Run tests with coverage
npm run test:coverage

# Lint code
npm run lint

# Format code
npm run format
```

## License

MIT License
