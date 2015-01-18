using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookDndWinForms
{
    public partial class Form1 : Form
    {
        private LocalDropTarget myDropTarget = null;

        public Form1()
        {
            InitializeComponent();
        }

        public IOleDropTarget GetDropTarget()
        {
            this.myDropTarget = new LocalDropTarget();
            return myDropTarget;
        }

        class LocalDropTarget : IOleDropTarget
        {
            public void OnDragDrop(System.Windows.DataObject d)
            {
                Trace.WriteLine("OnDragDrop");

                var formats = d.GetFormats();

                foreach (var format in formats)
                    Trace.WriteLine(format);

                var obj = d;
                var data = obj.GetData("RenPrivateMessages");

                if (data is MemoryStream)
                {
                    BinaryReader reader = new BinaryReader(data as MemoryStream);
                    int folderIdLength = reader.ReadInt32();
                }
            }

            public int OleDragEnter(object pDataObj, int grfKeyState, long pt, ref int pdwEffect)
            {
                Trace.WriteLine("OleDragEnter");
                Marshal.FinalReleaseComObject(pDataObj);
                return 0;
            }

            public int OleDragOver(int grfKeyState, long pt, ref int pdwEffect)
            {
                Trace.WriteLine("OleDragOver");
                return 0;
            }

            public int OleDragLeave()
            {
                Trace.WriteLine("OleDragEnter");
                return 0;
            }

            public int OleDrop(object pDataObj, int grfKeyState, long pt, ref int pdwEffect)
            {
                Trace.WriteLine("OleDrop");
                System.Windows.DataObject data = new System.Windows.DataObject(pDataObj);

                OnDragDrop(data);
                Marshal.FinalReleaseComObject(pDataObj);
                return 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var res = RegisterDragDrop(this.panel1.Handle, GetDropTarget());
        }

        //Native imports
        [DllImport("ole32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        public static extern int RegisterDragDrop(IntPtr hwnd, IOleDropTarget target);

        [ComImport(), Guid("00000122-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleDropTarget
        {
            [PreserveSig]
            int OleDragEnter(
                [In, MarshalAs(UnmanagedType.Interface)]
                object pDataObj,
                [In, MarshalAs(UnmanagedType.U4)]
                int grfKeyState,
                [In, MarshalAs(UnmanagedType.U8)]
                long pt,
                [In, Out]
                ref int pdwEffect);

            [PreserveSig]
            int OleDragOver(
                [In, MarshalAs(UnmanagedType.U4)]
                int grfKeyState,
                [In, MarshalAs(UnmanagedType.U8)]
                long pt,
                [In, Out]
                ref int pdwEffect);

            [PreserveSig]
            int OleDragLeave();

            [PreserveSig]
            int OleDrop(
                [In, MarshalAs(UnmanagedType.Interface)]
                object pDataObj,
                [In, MarshalAs(UnmanagedType.U4)]
                int grfKeyState,
                [In, MarshalAs(UnmanagedType.U8)]
                long pt,
                [In, Out]
                ref int pdwEffect);
        }

    }
}
