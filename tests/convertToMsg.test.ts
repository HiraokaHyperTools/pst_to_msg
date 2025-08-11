import { convertToMsg } from '../src/index';
import { openPstFile, PSTFolder } from '@hiraokahypertools/pst-extractor';
import { writeFile, mkdir } from 'node:fs/promises';

function normFileName(fileName: string): string {
  return fileName.replace(/[\\\/\:\*\?\"\<\>\|]/g, '_');
}

describe('convertToMsg', () => {
  describe('batch conversion', () => {

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

    it('200 recipients.pst', async () => await extractPstFileTo('pst/200 recipients.pst', 'temp/200 recipients'));
    it('alpha-beta-gamma-delta.pst', async () => await extractPstFileTo('pst/alpha-beta-gamma-delta.pst', 'temp/alpha-beta-gamma-delta'));
    it('attachAndInline.pst', async () => await extractPstFileTo('pst/attachAndInline.pst', 'temp/attachAndInline'));
    it('contacts.pst', async () => await extractPstFileTo('pst/contacts.pst', 'temp/contacts'));
    it('contacts97-2002.pst', async () => await extractPstFileTo('pst/contacts97-2002.pst', 'temp/contacts97-2002'));
    it('manyFiles.pst', async () => await extractPstFileTo('pst/manyFiles.pst', 'temp/manyFiles'));
    it('msgInMsg.pst', async () => await extractPstFileTo('pst/msgInMsg.pst', 'temp/msgInMsg'));
    it('msgInMsgInMsg.pst', async () => await extractPstFileTo('pst/msgInMsgInMsg.pst', 'temp/msgInMsgInMsg'));
    it('nonUnicodeCP932.pst', async () => await extractPstFileTo('pst/nonUnicodeCP932.pst', 'temp/nonUnicodeCP932'));
    it('Outlook97-2002.pst', async () => await extractPstFileTo('pst/Outlook97-2002.pst', 'temp/Outlook97-2002'));
    it('Outlook2003.pst', async () => await extractPstFileTo('pst/Outlook2003.pst', 'temp/Outlook2003'));
    it('recipientNameHasUnicodeChars.pst', async () => await extractPstFileTo('pst/recipientNameHasUnicodeChars.pst', 'temp/recipientNameHasUnicodeChars'));
    it('simple.pst', async () => await extractPstFileTo('pst/simple.pst', 'temp/simple'));
    it('unicodeAttachmentFilename.pst', async () => await extractPstFileTo('pst/unicodeAttachmentFilename.pst', 'temp/unicodeAttachmentFilename'));
  });
});
