using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0, ext_cm_Startup = 1, ext_cm_External = 2,
        ext_cm_CommandLine = 3, ext_cm_Solution = 4, ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2, ext_dm_SolutionClosed = 3
    }

    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application,
            ext_ConnectMode connectMode, [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            ref Array custom);
        void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }
}

namespace Microsoft.Office.Core
{
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    [ComImport]
    [Guid("000C03A7-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonUI
    {
        void Invalidate();
        void InvalidateControl([MarshalAs(UnmanagedType.BStr)] string controlId);
        void Refresh();
        void ActivateTab([MarshalAs(UnmanagedType.BStr)] string controlId);
        void ActivateTabMso([MarshalAs(UnmanagedType.BStr)] string controlId);
        void ActivateTabQ([MarshalAs(UnmanagedType.BStr)] string controlId,
            [MarshalAs(UnmanagedType.BStr)] string namespaceName);
    }

    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonControl
    {
        string Id { [return: MarshalAs(UnmanagedType.BStr)] get; }
        object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
        string Tag { [return: MarshalAs(UnmanagedType.BStr)] get; }
    }
}
