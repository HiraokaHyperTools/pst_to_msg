# pst_to_msg

[![npm](https://img.shields.io/npm/v/@hiraokahypertools/pst_to_msg)](https://www.npmjs.com/package/@hiraokahypertools/pst_to_msg)

Convert a message in Outlook PST file into MSG format

Links: [typedoc documentation](https://hiraokahypertools.github.io/pst_to_msg/typedoc/)

## Export msg files in a batch work

Traverse the entire pst file, and extract the found messages into file system folders.

```ts
import { convertToMsg } from '../src/index';
import { openPstFile, PSTFolder } from '@hiraokahypertools/pst-extractor';
import { writeFile, mkdir } from 'node:fs/promises';

const extractPstFileTo = async (pstFilePath: string, outDirRoot: string): Promise<void> => {
  const pstFile = await openPstFile(pstFilePath, { ansiEncoding: "cp932", });
  try {
    async function walk(folder: PSTFolder, outDir: string): Promise<void> {
      {
        const subFolderCount = await folder.getSubFolderCount();
        for (let i = 0; i < subFolderCount; i++) {
          const subFolder = await folder.getSubFolder(i);
          await walk(subFolder, `${outDir}${normFileName(subFolder.displayName)}/`);
        }
      }
      {
        const emailCount = await folder.getEmailCount();
        for (let i = 0; i < emailCount; i++) {
          const message = await folder.getEmail(i);
          const msg = await convertToMsg(message, pstFile);
          await mkdir(outDir, { recursive: true });
          const saveTo = `${outDir}${normFileName(message.subject)}.msg`;
          await writeFile(saveTo, msg);
        }
      }
    }

    const rootFolder = (await pstFile.getRootFolder())!;
    await walk(rootFolder, outDirRoot ? outDirRoot + "/" : "");
  } finally {
    pstFile.close();
  }
};

await extractPstFileTo('pst/manyFiles.pst', 'temp/manyFiles');
```
