Reverse Engineered Outlook Drag and Drop (Clipboard) OLE Protocol
===================

Introduction
-------------------

If you want to implement drag and drop support from Microsoft Outlook to your application, there are already some good [resources](http://www.codeproject.com/Articles/28209/Outlook-Drag-and-Drop-in-C) out there.

These only describe how you can get the temporary .msg-File Outlook creates in a temporary folder. If you want information about the original Outlook message (like EntryID, StoreID, etc.), this won't help you.

In the process of buiding our Outlook addin, we came across this problem, as we need to identify the message in Outlook. There are some workarounds using the current Outlook selection, but we thought there might be another way!

It turns out, there is. Outlook also provides some formats called "RenPrivateMessages", "RenPrivateLatestMessages" etc. You can view all of that data using [ClipSpy](http://www.codeproject.com/Articles/168/ClipSpy), which we also used to reverse engineere this stuff. You can find an example WPF project in this repository.

The Format
------------------

It's actually a quite simple format, even though there are some bytes missing, the most interesting stuff is easy to get.
Once you get the byte stream of "RenPrivateMessages", it can be read using the following parser:

<table>
  <tr>
    <th>Byte</th><th>Length</th><th>Value</th>
  </tr>
  <tr>
    <td>0</td><td>4 Bytes</td><td>FolderId length</td>
  </tr>
  <tr>
    <td>5</td><td>Length given by previous value</td><td>The MAPI ParentFolderId of the item</td>
  </tr>
</table>
