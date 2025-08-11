import { PSTFile, PSTMessage } from "@hiraokahypertools/pst-extractor";
import { burn, Entry } from "@kenjiuno/msgreader/lib/Burner";
import { TypeEnum } from "@kenjiuno/msgreader/lib/Reader";
import { toHex2, toHex4 } from "./utils";
import { PUSubNode } from "@hiraokahypertools/pst-extractor/dist/PUSubNode";
import { RawProperty } from "@hiraokahypertools/pst-extractor/dist/RawProperty";

type Prop1 = {
  tagLo: number;
  tagHi: number;
  mandatory: boolean;
  readable: boolean;
  writeable: boolean;
  value: Uint8Array;
};

export async function convertToMsg(
  message: PSTMessage,
  file: PSTFile,
): Promise<Uint8Array> {
  const entries: Entry[] = [
    {
      name: "Root Entry",
      type: TypeEnum.ROOT,
      children: [],
      length: 0,
    },
  ];

  await convertMessage(
    entries,
    0,
    (await message.requestAccessToUserSubNode())!,
    true,
    file.getStoreSupportMask(),
  );

  {
    const dirIndex = addDir(entries, 0, "__nameid_version1.0");
    const node = (await file.requestAccessToUserNode(97))!;
    const subNode = (await node.getSubNode())!;
    const pc = await subNode.extractAsPropertyContext();
    await convertNameId(
      entries,
      dirIndex,
      pc.properties,
      subNode,
      pc.resolveHeap,
    );
  }

  return burn(entries);
}

async function convertMessage(
  entries: Entry[],
  parentIndex: number,
  subNode: PUSubNode,
  isRoot: boolean,
  storeSupportMask: number | undefined,
): Promise<void> {
  // `__properties_version1.0`
  // `__substg1.0_0E1D001F`

  // `__nameid_version1.0`
  // - `__substg1.0_100A0102`

  // `__recip_version1.0_#00000000` ← Upper hexa 8 digits
  // - `__properties_version1.0`
  // - `__substg1.0_0FF60102`

  // `__attach_version1.0_#00000000`
  // - `__properties_version1.0`
  // - `__substg1.0_0FF90102`
  // - `__substg1.0_3701000D`
  //   - `__recip_version1.0_#00000000`
  //   - `__properties_version1.0`
  //   - `__substg1.0_0C1A001F`

  let recipientsCount = 0;

  const recipientsSubNode = await subNode.getSubNodeOf(0x692);
  if (recipientsSubNode) {
    const tc = await recipientsSubNode.extractAsTableContext();
    recipientsCount = tc.numRows;

    for (let y = 0; y < recipientsCount; y++) {
      const dirIndex = addDir(
        entries,
        parentIndex,
        `__recip_version1.0_#${toHex4(y, true)}`,
      );
      await convertGeneric(
        entries,
        dirIndex,
        await tc.getRow(y),
        recipientsSubNode,
        tc.resolveHeap,
        y,
      );
    }
  }

  let attachmentsCount = 0;

  const attachmentsSubNode = await subNode.getSubNodeOf(0x671);
  if (attachmentsSubNode) {
    const tc = await attachmentsSubNode.extractAsTableContext();
    attachmentsCount = tc.numRows;

    for (let y = 0; y < attachmentsCount; y++) {
      const dirIndex = addDir(
        entries,
        parentIndex,
        `__attach_version1.0_#${toHex4(y, true)}`,
      );
      const limitedProperties = await tc.getRow(y);
      const ltpRowId = limitedProperties.find(
        (it) => it.key == 0x67f2 && it.type == 0x0003,
      );
      if (ltpRowId) {
        const rowId = new DataView(ltpRowId.value).getUint32(0, true);
        const attachmentSubNode = await subNode.getSubNodeOf(rowId);
        if (attachmentSubNode) {
          const pc = await attachmentSubNode.extractAsPropertyContext();
          await convertAtt(
            entries,
            dirIndex,
            pc.properties,
            attachmentSubNode,
            pc.resolveHeap,
            y,
          );
        } else {
          throw new Error(
            `Attachment sub-node with row ID ${rowId} not found.`,
          );
        }
      } else {
        throw new Error(`LTP row ID ${ltpRowId} not found.`);
      }
    }
  }

  {
    const props1: Prop1[] = [];

    const pc = await subNode.extractAsPropertyContext();
    for (let y = 0; y < pc.properties.length; y++) {
      const pair = await convertProperty(
        entries,
        parentIndex,
        pc.properties[y],
        subNode,
        pc.resolveHeap,
        false,
      );
      props1.push(pair.prop1);
    }

    if (isRoot) {
      // At least 3 properties are mandatory for a message root:
      // - PidTagStoreSupportMask
      // - PidTagMessageFlags (also included in pst)
      // - PidTagMessageClass (also included in pst)

      const arrayBuffer = new ArrayBuffer(8);
      const view = new DataView(arrayBuffer);

      // from "C:\Program Files (x86)\Windows Kits\10\Include\10.0.20348.0\um\WabDefs.h"
      // from https://github.com/microsoft/mfcmapi/blob/7e3111d899638da0aa594c06f5a00f980f70b0b8/core/mapi/extraPropTags.h#L1056-L1059

      if (typeof storeSupportMask === "number") {
        const STORE_ENTRYID_UNIQUE = 0x00000001;
        const STORE_READONLY = 0x00000002;
        const STORE_SEARCH_OK = 0x00000004;
        const STORE_MODIFY_OK = 0x00000008;
        const STORE_CREATE_OK = 0x00000010;
        const STORE_ATTACH_OK = 0x00000020;
        const STORE_OLE_OK = 0x00000040;
        const STORE_SUBMIT_OK = 0x00000080;
        const STORE_NOTIFY_OK = 0x00000100;
        const STORE_MV_PROPS_OK = 0x00000200;
        const STORE_CATEGORIZE_OK = 0x00000400;
        const STORE_RTF_OK = 0x00000800;
        const STORE_RESTRICTION_OK = 0x00001000;
        const STORE_SORT_OK = 0x00002000;
        const STORE_HAS_SEARCHES = 0x01000000;
        const STORE_ANSI_OK = 0x00020000;
        const STORE_HTML_OK = 0x00010000;
        const STORE_ITEMPROC = 0x00200000;
        const STORE_LOCALSTORE = 0x00080000;
        const STORE_PUBLIC_FOLDERS = 0x00004000;
        const STORE_PUSHER_OK = 0x00800000;
        const STORE_RULES_OK = 0x10000000;
        const STORE_UNCOMPRESSED_RTF = 0x00008000;
        const STORE_UNICODE_OK = 0x00040000;

        // // 0x40E79
        // view.setUint32(0, 0
        //   | STORE_ENTRYID_UNIQUE | STORE_MODIFY_OK
        //   | STORE_CREATE_OK | STORE_ATTACH_OK | STORE_OLE_OK | STORE_MV_PROPS_OK | STORE_CATEGORIZE_OK | STORE_RTF_OK
        //   | STORE_UNICODE_OK,
        //   true
        // );

        view.setUint32(0, storeSupportMask, true);

        // PidTagStoreSupportMask
        props1.push({
          tagLo: 0x0003,
          tagHi: 0x340d,
          mandatory: false,
          readable: true,
          writeable: false,
          value: new Uint8Array(arrayBuffer, 0, 4),
        });
      }

      // {
      //   const MSGFLAG_READ = 0x00000001;
      //   const MSGFLAG_UNMODIFIED = 0x00000002;
      //   const MSGFLAG_SUBMIT = 0x00000004;
      //   const MSGFLAG_UNSENT = 0x00000008;
      //   const MSGFLAG_HASATTACH = 0x00000010;
      //   const MSGFLAG_FROMME = 0x00000020;
      //   const MSGFLAG_ASSOCIATED = 0x00000040;
      //   const MSGFLAG_RESEND = 0x00000080;
      //   const MSGFLAG_RN_PENDING = 0x00000100;
      //   const MSGFLAG_NRN_PENDING = 0x00000200;
      //   const MSGFLAG_ORIGIN_X400 = 0x00001000;
      //   const MSGFLAG_ORIGIN_INTERNET = 0x00002000;
      //   const MSGFLAG_ORIGIN_MISC_EXT = 0x00008000;
      //   const MSGFLAG_OUTLOOK_NON_EMS_XP = 0x00010000;

      //   view.setUint32(4, MSGFLAG_READ | MSGFLAG_UNSENT, true);

      //   // PidTagMessageFlags
      //   props1.push({
      //     tagLo: 0x0003,
      //     tagHi: 0x0E07,
      //     mandatory: false,
      //     readable: true,
      //     writeable: false,
      //     value: new Uint8Array(arrayBuffer, 4, 4),
      //   });
      // }
    }

    addFile(
      entries,
      parentIndex,
      "__properties_version1.0",
      (isRoot ? burnPropertyStreamHeader32 : burnPropertyStreamHeader24)(
        recipientsCount,
        attachmentsCount,
        props1,
      ),
    );
  }
}

async function convertGeneric(
  entries: Entry[],
  parentIndex: number,
  properties: RawProperty[],
  subNode: PUSubNode,
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>,
  rowIndex: number,
): Promise<void> {
  const props1: Prop1[] = [];

  for (let y = 0; y < properties.length; y++) {
    const pair = await convertProperty(
      entries,
      parentIndex,
      properties[y],
      subNode,
      resolveHeap,
      false,
    );
    props1.push(pair.prop1);
  }

  {
    // At least 1 property are mandatory for contact root:
    // - PidTagRowid

    const arrayBuffer = new ArrayBuffer(4);
    const view = new DataView(arrayBuffer);
    view.setUint32(0, rowIndex, true);

    props1.push({
      tagLo: 0x0003,
      tagHi: 0x3000,
      mandatory: false,
      readable: true,
      writeable: true,
      value: new Uint8Array(arrayBuffer, 0, 4),
    });
  }

  addFile(
    entries,
    parentIndex,
    "__properties_version1.0",
    burnPropertyStreamHeader8(props1),
  );
}

async function convertAtt(
  entries: Entry[],
  parentIndex: number,
  properties: RawProperty[],
  subNode: PUSubNode,
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>,
  attachmentIndex: number,
): Promise<void> {
  const props1: Prop1[] = [];

  for (let y = 0; y < properties.length; y++) {
    const pair = await convertProperty(
      entries,
      parentIndex,
      properties[y],
      subNode,
      resolveHeap,
      false,
    );
    props1.push(pair.prop1);
  }

  {
    // At least 2 properties are mandatory for attachment root:
    // - PidTagAttachNumber
    // - PidTagAttachMethod (also included in pst)

    const arrayBuffer = new ArrayBuffer(4);
    const view = new DataView(arrayBuffer);
    view.setUint32(0, attachmentIndex, true);

    props1.push({
      tagLo: 0x0003,
      tagHi: 0x0e21,
      mandatory: false,
      readable: true,
      writeable: false,
      value: new Uint8Array(arrayBuffer, 0, 4),
    });
  }

  addFile(
    entries,
    parentIndex,
    "__properties_version1.0",
    burnPropertyStreamHeader8(props1),
  );
}

async function convertNameId(
  entries: Entry[],
  parentIndex: number,
  properties: RawProperty[],
  subNode: PUSubNode,
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>,
): Promise<void> {
  const props1: Prop1[] = [];

  for (let y = 0; y < properties.length; y++) {
    const pair = await convertProperty(
      entries,
      parentIndex,
      properties[y],
      subNode,
      resolveHeap,
      true,
    );
    props1.push(pair.prop1);
  }
}

function addFile(
  entries: Entry[],
  parentIndex: number,
  name: string,
  data: Uint8Array,
): void {
  const entry: Entry = {
    name,
    type: TypeEnum.DOCUMENT,
    length: data.length,
    binaryProvider: () => data,
  };
  const fileIndex = entries.length;
  entries.push(entry);
  entries[parentIndex].children = (entries[parentIndex].children || []).concat(
    fileIndex,
  );
}

function addDir(entries: Entry[], parentIndex: number, name: string): number {
  const entry: Entry = {
    name,
    type: TypeEnum.DIRECTORY,
    children: [],
    length: 0,
  };
  const newIndex = entries.length;
  entries.push(entry);
  entries[parentIndex].children = (entries[parentIndex].children || []).concat(
    newIndex,
  );
  return newIndex;
}

function burnPropertyStreamHeader32(
  recipientsCount: number,
  attachmentsCount: number,
  props: Prop1[],
): Uint8Array {
  const array = new Uint8Array(32 + 16 * props.length);
  const view = new DataView(array.buffer);
  view.setUint32(0x08, recipientsCount, true); // int NextRecipientID
  view.setUint32(0x0c, attachmentsCount, true); // int NextAttachmentID
  view.setUint32(0x10, recipientsCount, true); // int RecipientCount
  view.setUint32(0x14, attachmentsCount, true); // int AttachmentCount
  for (let y = 0; y < props.length; y++) {
    const top = 32 + 16 * y;
    const one = props[y];
    view.setUint16(top + 0x0, one.tagLo, true);
    view.setUint16(top + 0x2, one.tagHi, true);
    view.setUint32(
      top + 0x4,
      (one.mandatory ? 1 : 0) |
      (one.readable ? 2 : 0) |
      (one.writeable ? 4 : 0),
      true,
    );
    if (8 < one.value.length) {
      throw new Error("Property value has to be at most 8 bytes long!");
    }
    array.set(one.value, top + 0x8);
  }
  return array;
}

function burnPropertyStreamHeader24(
  recipientsCount: number,
  attachmentsCount: number,
  props: Prop1[],
): Uint8Array {
  const array = new Uint8Array(24 + 16 * props.length);
  const view = new DataView(array.buffer);
  view.setUint32(0x08, recipientsCount, true); // int NextRecipientID
  view.setUint32(0x0c, attachmentsCount, true); // int NextAttachmentID
  view.setUint32(0x10, recipientsCount, true); // int RecipientCount
  view.setUint32(0x14, attachmentsCount, true); // int AttachmentCount
  for (let y = 0; y < props.length; y++) {
    const top = 24 + 16 * y;
    const one = props[y];
    view.setUint16(top + 0x0, one.tagLo, true);
    view.setUint16(top + 0x2, one.tagHi, true);
    view.setUint32(
      top + 0x4,
      (one.mandatory ? 1 : 0) |
      (one.readable ? 2 : 0) |
      (one.writeable ? 4 : 0),
      true,
    );
    if (8 < one.value.length) {
      throw new Error("Property value has to be at most 8 bytes long!");
    }
    array.set(one.value, top + 0x8);
  }
  return array;
}

function burnPropertyStreamHeader8(props: Prop1[]): Uint8Array {
  const array = new Uint8Array(8 + 16 * props.length);
  const view = new DataView(array.buffer);
  for (let y = 0; y < props.length; y++) {
    const top = 8 + 16 * y;
    const one = props[y];
    view.setUint16(top + 0x0, one.tagLo, true);
    view.setUint16(top + 0x2, one.tagHi, true);
    view.setUint32(
      top + 0x4,
      (one.mandatory ? 1 : 0) |
      (one.readable ? 2 : 0) |
      (one.writeable ? 4 : 0),
      true,
    );
    if (8 < one.value.length) {
      throw new Error("Property value has to be at most 8 bytes long!");
    }
    array.set(one.value, top + 0x8);
  }
  return array;
}

async function convertProperty(
  entries: Entry[],
  parentIndex: number,
  prop: RawProperty,
  subNode: PUSubNode,
  resolveHeap: (heap: number) => Promise<ArrayBuffer | undefined>,
  writeAllPropsToFile: boolean,
): Promise<{ prop1: Prop1 }> {
  let value = prop.value;

  // console.log("convertProperty", toHex2(prop.key), toHex2(prop.type), prop.value?.byteLength, prop.value);
  if (
    false ||
    prop.type === 0x000b ||
    prop.type === 0x0002 ||
    prop.type === 0x0004 ||
    prop.type === 0x0003
  ) {
    if (writeAllPropsToFile) {
      if (value) {
        addFile(
          entries,
          parentIndex,
          `__substg1.0_${toHex2(prop.key, true)}${toHex2(prop.type, true)}`,
          new Uint8Array(value),
        );
      }
    }
  } else if (prop.type === 0x000d) {
    // `__substg1.0_3701000D`
    const dirIndex = addDir(
      entries,
      parentIndex,
      `__substg1.0_${toHex2(prop.key, true)}000D`,
    );
    const hnid = new DataView(value).getUint32(0, true);
    const bytes = await resolveHeap(hnid);
    if (bytes !== undefined) {
      const view = new DataView(bytes);
      const subNodeId = view.getUint32(0, true);

      const subSubNode = await subNode.getSubNodeOf(subNodeId);
      if (subSubNode) {
        await convertMessage(entries, dirIndex, subSubNode, false, undefined);

        value = new ArrayBuffer(8);
        const view = new DataView(value);
        view.setUint32(0, 0, true);
        view.setUint32(4, 1, true); // ATTACH_EMBEDDED_MSG (0x00000005) → 0x01
      }
    }
  } else {
    if (value && 4 <= value.byteLength) {
      const hnid = new DataView(value).getUint32(0, true);
      const bytes = await resolveHeap(hnid);
      if (bytes !== undefined) {
        const apply = (actualData: ArrayBuffer): void => {
          addFile(
            entries,
            parentIndex,
            `__substg1.0_${toHex2(prop.key, true)}${toHex2(prop.type, true)}`,
            new Uint8Array(actualData),
          );
        };
        if (false) {
        } else if (prop.type === 0x1e) {
          const altBytes = fixAnsiString(bytes);
          value = new ArrayBuffer(8);
          const view = new DataView(value);
          view.setUint32(0, altBytes.byteLength + 1, true);
          apply(altBytes);
        } else if (prop.type === 0x1f) {
          const altBytes = fixUnicodeString(bytes);
          value = new ArrayBuffer(8);
          const view = new DataView(value);
          view.setUint32(0, altBytes.byteLength + 2, true);
          apply(altBytes);
        } else {
          value = new ArrayBuffer(8);
          const view = new DataView(value);
          view.setUint32(0, bytes.byteLength, true);
          apply(bytes);
        }
      } else {
        addFile(
          entries,
          parentIndex,
          `__substg1.0_${toHex2(prop.key, true)}${toHex2(prop.type, true)}`,
          new Uint8Array(0),
        );
      }
    } else {
      addFile(
        entries,
        parentIndex,
        `__substg1.0_${toHex2(prop.key, true)}${toHex2(prop.type, true)}`,
        new Uint8Array(0),
      );
    }
  }

  return {
    prop1: {
      tagLo: prop.type,
      tagHi: prop.key,
      mandatory: false,
      readable: true,
      writeable: true,
      value: value ? new Uint8Array(value) : new Uint8Array(0),
    },
  };
}

function fixUnicodeString(bytes: ArrayBuffer): ArrayBuffer {
  if (4 <= bytes.byteLength) {
    const view = new Uint8Array(bytes);
    return view[0] === 1 && view[1] === 0 && view[2] === 1 && view[3] === 0
      ? bytes.slice(4)
      : bytes;
  } else {
    return bytes;
  }
}

function fixAnsiString(bytes: ArrayBuffer): ArrayBuffer {
  if (2 <= bytes.byteLength) {
    const view = new Uint8Array(bytes);
    return view[0] === 1 && view[1] === 1 ? bytes.slice(2) : bytes;
  } else {
    return bytes;
  }
}
