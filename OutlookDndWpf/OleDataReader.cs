// --------------------------------------------------------------------------
// Licensed under MIT License.
//
// Outlook DnD Data Reader
// 
// File     : OleDataReader.cs
// Author   : Tobias Viehweger <tobias.viehweger@yasoon.com / @mnkypete>
//
// -------------------------------------------------------------------------- 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookDndWpf
{
    internal class OleDataReader
    {
        private MemoryStream stream;

        public OleDataReader(MemoryStream inStream)
        {
            this.stream = inStream;
        }

        public OleOutlookData ReadOutlookData()
        {
            BinaryReader reader = new BinaryReader(this.stream);

            //1. First 4 bytes are the length of the FolderId (In bytes)
            // Note: These are possibly uint? We don't expect it to be that long nevertheless..
            int folderIdLength = reader.ReadInt32();

            //2. Read FolderId
            byte[] folderId = reader.ReadBytes(folderIdLength);
            string folderIdHex = ByteArrayToString(folderId);

            //3. Next 4 bytes are the StoreId length (In bytes)
            int storeIdLength = reader.ReadInt32();

            //4. Read StoreId
            byte[] storeId = reader.ReadBytes(storeIdLength);
            string storeIdHex = ByteArrayToString(storeId);

            //5. There are now some bytes which are not identified yet..
            reader.ReadBytes(4);
            reader.ReadBytes(4);
            reader.ReadBytes(4); // <== These appear to be folder dependent somehow..

            //6. Read items count, again, we assume int instead of uint because that much items
            //   => Other problems =)
            int itemCount = reader.ReadInt32();

            OleOutlookItemData[] items = new OleOutlookItemData[itemCount];

            for (int i = 0; i < itemCount; i++)
            {
                //First 4 bytes, represent the MAPI property 0x8014 ("SideEffects" in OlSpy)
                int sideEffects = reader.ReadInt32();

                //Next byte tells us the length of the message class string (i.e. IPM.Note)
                byte classLength = reader.ReadByte();

                //Now, read type
                string messageClass = Encoding.ASCII.GetString(reader.ReadBytes(classLength));

                //Next, read the unicode char (!) count of the subject 
                // Note: It seems that Outlook limits this to 255, cross reference mail spec sometime..
                byte subjectLength = reader.ReadByte();

                //Read the subject, note that this is unicode, so we need to read 2 bytes per char!
                string subject = Encoding.Unicode.GetString(reader.ReadBytes(subjectLength * 2));

                //Next up: EntryID including it's length (same as for store + folder)
                int entryIdLength = reader.ReadInt32();
                byte[] entryId = reader.ReadBytes(entryIdLength);
                string entryIdHex = ByteArrayToString(entryId);

                //Now the SearchKey MAPI property of the item
                int searchKeyLength = reader.ReadInt32();
                byte[] searchKey = reader.ReadBytes(searchKeyLength);
                string searchKeyHex = ByteArrayToString(searchKey);

                //Some more stuff which is not quite clear, the next 4 bytes seem to be always
                // => E0 80 E9 5A
                reader.ReadBytes(4);

                //The next 24 byte are some more flags which are not worked out yet, afterwards
                // the next item begins
                reader.ReadBytes(24);

                items[i] = new OleOutlookItemData { 
                    EntryId = entryIdHex, 
                    MessageClass = messageClass, 
                    SearchKey = searchKeyHex, 
                    Subject = subject 
                };
            }

            OleOutlookData data = new OleOutlookData();
            data.StoreId = storeIdHex;
            data.FolderId = folderIdHex;
            data.Items = items;

            return data;
        }

        private string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2}", b);
            return hex.ToString();
        }
    }
}
