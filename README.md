Reverse Engineered Outlook Drag and Drop (Clipboard) OLE Protocol
===================

Introduction
-------------------

If you want to implement drag and drop support from Microsoft Outlook to your application, there are already some good [resources](http://www.codeproject.com/Articles/28209/Outlook-Drag-and-Drop-in-C) out there.

These only describe how you can get the temporary .msg-File Outlook creates in a temporary folder. If you want information about the original Outlook message (like EntryID, StoreID, etc.), this won't help you.

In the process of buiding our Outlook addin, we came across this problem, as we need to identify the message in Outlook. There are some workarounds using the current Outlook selection, but we thought there might be another way!

It turns out, there is. Outlook also provides some formats called "RenPrivateMessages", "RenPrivateLatestMessages" etc. You can view all of that data using [ClipSpy](http://www.codeproject.com/Articles/168/ClipSpy), which we also used to reverse engineere this stuff. You can find an example WPF project in this repository.

Note: This was only tested in Outlook 2013 yet, 2010 will be next.

The Format
------------------

It's actually a quite simple format, even though there are some bytes missing, the most interesting stuff is easy to get.
Once you get the byte stream of "RenPrivateMessages", it can be read using the following parser:

<table>
  <tr>
    <th>Length in Bytes</th><th>Type</th><th>Value</th>
  </tr>
  <tr>
    <td>4 Bytes</td><td>int (possibly uint)</td><td>FolderId length</td>
  </tr>
  <tr>
    <td>Length given by previous value</td><td>binary</td><td>The MAPI ParentFolderId of the item</td>
  </tr>
  <tr>
    <td>4 Bytes</td><td>int</td><td>StoreId length</td>
  </tr>
  <tr>
    <td>Length given by previous value</td><td>binary</td><td>The MAPI StoreId of the item</td>
  </tr>
  <tr>
    <td>4 Byte</td><td>??</td><td>Unknown</td>
  </tr>
  <tr>
    <td>4 Byte</td><td>??</td><td>Unknown</td>
  </tr>
  <tr>
    <td>4 Byte</td><td>??</td><td>Unknown, but seems to be folder dependent</td>
  </tr>
  <tr>
    <td>4 Byte</td><td>int</td><td>Number of Items</td>
  </tr>
  <tr>
    <td>Loop for itemCount</td><td>---</td><td>---</td>
  </tr>
  <tr>
    <td>4 Byte</td><td>int</td><td>Represent the MAPI property 0x8014 ("SideEffects" in OlSpy)</td>
  </tr>
  <tr>
    <td>1 Byte</td><td>byte</td><td>Length of MessageClass</td>
  </tr>
  <tr>
    <td>Length given by previous value</td><td>ASCII</td><td>The MessageClass (e.g. IPM.Task) of the item</td>
  </tr>
  <tr>
    <td>1 Byte</td><td>byte</td><td>Number of Unicode chars of Subject</td>
  </tr>
  <tr>
    <td>Number of Unicode chars * 2</td><td>Unicode</td><td>The subject of the item</td>
  </tr>
  <tr>
    <td>4 Bytes</td><td>int</td><td>EntryId length</td>
  </tr>
  <tr>
    <td>Length given by previous value</td><td>binary</td><td>The MAPI EntryId of the item</td>
  </tr>
  <tr>
    <td>4 Bytes</td><td>int</td><td>SearchKey length</td>
  </tr>
  <tr>
   <td>Length given by previous value</td><td>binary</td><td>The MAPI SearchKey of the item</td>
  </tr>
  <tr>
    <td>4 Bytes</td><td>??</td><td>Unknown, seems to be always E0 80 E9 5A</td>
  </tr>
  <tr>
    <td>24 Bytes</td><td>??</td><td>Unknown, seemingly some flags</td>
  </tr>
  <tr>
    <td>Next item </td><td>---</td><td>---</td>
  </tr>
</table>
