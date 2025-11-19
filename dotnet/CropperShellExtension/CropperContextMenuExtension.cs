using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IBindCtx = System.Runtime.InteropServices.ComTypes.IBindCtx;

namespace CropperShellExtension;

[ComVisible(true)]
[Guid(ClassGuid)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class CropperContextMenuExtension : IContextMenu, IShellExtInit, IExplorerCommand
{
    private const string ClassGuid = "7F56439E-2130-4115-A27A-8D562049B848";
    private const string MenuText = "Crop with PyCropper";
    private const string HelpText = "Crop selected file with PyCropper";
    private const string CropperExecutable = "cropper.exe";
    private const string ContextMenuKey = @"*\shellex\ContextMenuHandlers\PyCropper";
    private const string ExplorerCommandKey = @"*\shell\PyCropper";
    private const string Verb = "pycropper";

    private string[] _selectedFiles = Array.Empty<string>();

    public void Initialize(IntPtr pidlFolder, IDataObject? dataObject, IntPtr hKeyProgID)
    {
        _selectedFiles = ExtractFilePaths(dataObject);
    }

    public int QueryContextMenu(IntPtr hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags)
    {
        if (_selectedFiles.Length == 0 || (uFlags & NativeMethods.CMF_DEFAULTONLY) != 0)
        {
            return NativeMethods.MakeHResultSuccess(0);
        }

        if (!NativeMethods.InsertMenu(hMenu, indexMenu, NativeMethods.MF_BYPOSITION, idCmdFirst, MenuText))
        {
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Failed to insert the PyCropper context menu item.");
        }

        return NativeMethods.MakeHResultSuccess(1);
    }

    public void InvokeCommand(IntPtr pici)
    {
        var cbSize = Marshal.ReadInt32(pici);
        if (cbSize >= Marshal.SizeOf<NativeMethods.CMINVOKECOMMANDINFOEX>())
        {
            var infoEx = Marshal.PtrToStructure<NativeMethods.CMINVOKECOMMANDINFOEX>(pici);
            if (IsOurCommand(infoEx))
            {
                LaunchCropper();
            }

            return;
        }

        var info = Marshal.PtrToStructure<NativeMethods.CMINVOKECOMMANDINFO>(pici);
        if (IsOurCommand(info))
        {
            LaunchCropper();
        }
    }

    public void GetCommandString(UIntPtr idCmd, uint uFlags, IntPtr pReserved, StringBuilder? pszName, uint cchMax)
    {
        if (idCmd != UIntPtr.Zero || pszName is null || cchMax == 0)
        {
            return;
        }

        var text = (uFlags & NativeMethods.GCS_HELPTEXTW) == NativeMethods.GCS_HELPTEXTW ? HelpText : Verb;

        if (text.Length + 1 > cchMax)
        {
            text = text[..((int)cchMax - 1)];
        }

        pszName.Clear();
        _ = pszName.Append(text);
    }

    void IExplorerCommand.GetTitle(IShellItemArray? psiItemArray, out IntPtr ppszName)
    {
        ppszName = AllocateCoTaskMem(MenuText);
    }

    void IExplorerCommand.GetIcon(IShellItemArray? psiItemArray, out IntPtr ppszIcon)
    {
        ppszIcon = IntPtr.Zero;
    }

    void IExplorerCommand.GetToolTip(IShellItemArray? psiItemArray, out IntPtr ppszInfotip)
    {
        ppszInfotip = AllocateCoTaskMem(HelpText);
    }

    void IExplorerCommand.GetCanonicalName(out Guid pguidCommandName)
    {
        pguidCommandName = new Guid(ClassGuid);
    }

    void IExplorerCommand.GetState(IShellItemArray? psiItemArray, bool fOkToBeSlow, out EXPCMDSTATE pCmdState)
    {
        pCmdState = HasSelection(psiItemArray) ? EXPCMDSTATE.ECS_ENABLED : EXPCMDSTATE.ECS_HIDDEN;
    }

    void IExplorerCommand.Invoke(IShellItemArray? psiItemArray, IBindCtx? pbc)
    {
        var selection = ExtractFilePaths(psiItemArray);
        LaunchCropper(selection);
    }

    void IExplorerCommand.GetFlags(out EXPCMDFLAGS pFlags)
    {
        pFlags = EXPCMDFLAGS.ECF_DEFAULT | EXPCMDFLAGS.ECF_ALWAYS_SHOW;
    }

    void IExplorerCommand.EnumSubCommands(out IEnumExplorerCommand? ppEnum)
    {
        ppEnum = null;
    }

    private static IntPtr AllocateCoTaskMem(string text)
    {
        return Marshal.StringToCoTaskMemUni(text ?? string.Empty);
    }

    private static bool HasSelection(IShellItemArray? items)
    {
        if (items is null)
        {
            return false;
        }

        items.GetCount(out uint count);
        return count > 0;
    }

    private static string[] ExtractFilePaths(IShellItemArray? items)
    {
        if (items is null)
        {
            return Array.Empty<string>();
        }

        items.GetCount(out uint count);
        if (count == 0)
        {
            return Array.Empty<string>();
        }

        var results = new List<string>((int)count);

        for (uint i = 0; i < count; i++)
        {
            items.GetItemAt(i, out var shellItem);
            try
            {
                var path = GetFilePath(shellItem);
                if (!string.IsNullOrWhiteSpace(path))
                {
                    results.Add(path);
                }
            }
            finally
            {
                if (shellItem is not null)
                {
                    Marshal.ReleaseComObject(shellItem);
                }
            }
        }

        return results.ToArray();
    }

    private static string? GetFilePath(IShellItem item)
    {
        item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out var pszString);
        if (pszString == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            return Marshal.PtrToStringUni(pszString);
        }
        finally
        {
            Marshal.FreeCoTaskMem(pszString);
        }
    }

    [ComRegisterFunction]
    public static void Register(Type _)
    {
        using var shellExtensionKey = Registry.ClassesRoot.CreateSubKey(ContextMenuKey);
        shellExtensionKey?.SetValue(string.Empty, $"{{{ClassGuid}}}");

        using var explorerCommandKey = Registry.ClassesRoot.CreateSubKey(ExplorerCommandKey);
        if (explorerCommandKey is not null)
        {
            explorerCommandKey.SetValue("MUIVerb", MenuText);
            explorerCommandKey.SetValue("ExplorerCommandHandler", $"{{{ClassGuid}}}");
        }
    }

    [ComUnregisterFunction]
    public static void Unregister(Type _)
    {
        Registry.ClassesRoot.DeleteSubKey(ContextMenuKey, throwOnMissingSubKey: false);
        try
        {
            Registry.ClassesRoot.DeleteSubKeyTree(ExplorerCommandKey, throwOnMissingSubKey: false);
        }
        catch (ArgumentException)
        {
        }
    }

    private static string[] ExtractFilePaths(IDataObject? dataObject)
    {
        if (dataObject is null)
        {
            return Array.Empty<string>();
        }

        var format = new FORMATETC
        {
            cfFormat = NativeMethods.CF_HDROP,
            dwAspect = DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = TYMED.TYMED_HGLOBAL
        };

        STGMEDIUM medium = default;

        try
        {
            dataObject.GetData(ref format, out medium);
            if (medium.tymed != TYMED.TYMED_HGLOBAL || medium.unionmember == IntPtr.Zero)
            {
                return Array.Empty<string>();
            }

            uint fileCount = NativeMethods.DragQueryFile(medium.unionmember, 0xFFFFFFFF, null, 0);
            if (fileCount == 0)
            {
                return Array.Empty<string>();
            }

            var results = new List<string>((int)fileCount);
            var buffer = new StringBuilder(NativeMethods.MaxPath);

            for (uint i = 0; i < fileCount; i++)
            {
                buffer.Clear();
                uint chars = NativeMethods.DragQueryFile(medium.unionmember, i, buffer, (uint)buffer.Capacity);
                if (chars == 0)
                {
                    continue;
                }

                var path = buffer.ToString();
                if (!string.IsNullOrWhiteSpace(path))
                {
                    results.Add(path);
                }
            }

            return results.ToArray();
        }
        catch (COMException)
        {
            return Array.Empty<string>();
        }
        finally
        {
            if (medium.tymed != TYMED.TYMED_NULL)
            {
                NativeMethods.ReleaseStgMedium(ref medium);
            }
        }
    }

    private static bool IsOurCommand(NativeMethods.CMINVOKECOMMANDINFO info)
    {
        if (NativeMethods.IsVerbIdentifier(info.lpVerb))
        {
            return NativeMethods.LowWord(info.lpVerb) == 0;
        }

        var verb = Marshal.PtrToStringAnsi(info.lpVerb);
        return string.Equals(verb, Verb, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsOurCommand(NativeMethods.CMINVOKECOMMANDINFOEX info)
    {
        if (NativeMethods.IsVerbIdentifier(info.lpVerb))
        {
            return NativeMethods.LowWord(info.lpVerb) == 0;
        }

        var verb = NativeMethods.ResolveVerb(info);
        return string.Equals(verb, Verb, StringComparison.OrdinalIgnoreCase);
    }

    private void LaunchCropper(IEnumerable<string>? selection = null)
    {
        var candidates = selection ?? _selectedFiles;
        var target = candidates.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(target))
        {
            return;
        }

        try
        {
            var info = new ProcessStartInfo
            {
                FileName = CropperExecutable,
                UseShellExecute = false
            };

            info.ArgumentList.Add(target);
            _ = Process.Start(info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Unable to launch '{CropperExecutable}'. Ensure it is installed and available on PATH.\n{ex.Message}",
                "PyCropper Shell Extension",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}

#region COM contracts and interop helpers

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("000214e4-0000-0000-C000-000000000046")]
internal interface IContextMenu
{
    [PreserveSig]
    int QueryContextMenu(IntPtr hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);

    void InvokeCommand(IntPtr pici);

    void GetCommandString(UIntPtr idCmd, uint uFlags, IntPtr pReserved, StringBuilder? pszName, uint cchMax);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("000214e8-0000-0000-C000-000000000046")]
internal interface IShellExtInit
{
    void Initialize(IntPtr pidlFolder, IDataObject? pdtobj, IntPtr hKeyProgID);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("a08ce4d0-fa25-44ab-b530-33a241387151")]
internal interface IExplorerCommand
{
    void GetTitle(IShellItemArray? psiItemArray, out IntPtr ppszName);
    void GetIcon(IShellItemArray? psiItemArray, out IntPtr ppszIcon);
    void GetToolTip(IShellItemArray? psiItemArray, out IntPtr ppszInfotip);
    void GetCanonicalName(out Guid pguidCommandName);
    void GetState(IShellItemArray? psiItemArray, bool fOkToBeSlow, out EXPCMDSTATE pCmdState);
    void Invoke(IShellItemArray? psiItemArray, IBindCtx? pbc);
    void GetFlags(out EXPCMDFLAGS pFlags);
    void EnumSubCommands(out IEnumExplorerCommand? ppEnum);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("a88826f8-186f-4987-aade-ea0cef8fbfe8")]
internal interface IEnumExplorerCommand
{
    void Next(uint celt, out IExplorerCommand? pUICommand, out uint pceltFetched);
    void Skip(uint celt);
    void Reset();
    void Clone(out IEnumExplorerCommand? ppenum);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("b63ea76d-1f85-456f-a19c-48159efa858b")]
internal interface IShellItemArray
{
    void BindToHandler(IntPtr pbc, ref Guid rbhid, ref Guid riid, out IntPtr ppvOut);
    void GetPropertyStore(int flags, ref Guid riid, out IntPtr ppv);
    void GetPropertyDescriptionList(ref PROPERTYKEY keyType, ref Guid riid, out IntPtr ppv);
    void GetAttributes(uint attribFlags, uint sfgaoMask, out uint psfgaoAttribs);
    void GetCount(out uint pdwNumItems);
    void GetItemAt(uint dwIndex, out IShellItem ppsi);
    void EnumItems(out IntPtr ppenumShellItems);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
internal interface IShellItem
{
    void BindToHandler(IntPtr pbc, ref Guid rbhid, ref Guid riid, out IntPtr ppvOut);
    void GetParent(out IShellItem ppsi);
    void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
    void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    void Compare(IShellItem psi, uint hint, out int piOrder);
}

internal enum SIGDN : uint
{
    SIGDN_FILESYSPATH = 0x80058000
}

[Flags]
internal enum EXPCMDFLAGS : uint
{
    ECF_DEFAULT = 0x00000000,
    ECF_HASSUBCOMMANDS = 0x00000001,
    ECF_HASSPLITBUTTON = 0x00000002,
    ECF_ISSEPARATOR = 0x00000004,
    ECF_HIDELABEL = 0x00000008,
    ECF_HASSUBMENU = 0x00000010,
    ECF_HIDDEN = 0x00000020,
    ECF_ALWAYS_SHOW = 0x00000040
}

[Flags]
internal enum EXPCMDSTATE : uint
{
    ECS_ENABLED = 0,
    ECS_DISABLED = 0x00000001,
    ECS_HIDDEN = 0x00000002,
    ECS_CHECKED = 0x00000008
}

#pragma warning disable CS0649
internal struct PROPERTYKEY
{
    public Guid fmtid;
    public uint pid;
}
#pragma warning restore CS0649

internal static class NativeMethods
{
    public const short CF_HDROP = 0x000F;
    public const uint CMF_DEFAULTONLY = 0x00000001;
    public const uint MF_BYPOSITION = 0x00000400;
    public const uint GCS_HELPTEXTW = 0x00000005;
    public const uint CMIC_MASK_UNICODE = 0x00004000;
    public const int MaxPath = 1024;

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool InsertMenu(IntPtr hMenu, uint uPosition, uint uFlags, uint uIDNewItem, string lpNewItem);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern uint DragQueryFile(IntPtr hDrop, uint iFile, StringBuilder? lpszFile, uint cch);

    [DllImport("ole32.dll")]
    public static extern void ReleaseStgMedium(ref STGMEDIUM pmedium);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CMINVOKECOMMANDINFO
    {
        public int cbSize;
        public int fMask;
        public IntPtr hwnd;
        public IntPtr lpVerb;
        public IntPtr lpParameters;
        public IntPtr lpDirectory;
        public int nShow;
        public int dwHotKey;
        public IntPtr hIcon;
        public IntPtr lpTitle;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CMINVOKECOMMANDINFOEX
    {
        public int cbSize;
        public int fMask;
        public IntPtr hwnd;
        public IntPtr lpVerb;
        public IntPtr lpParameters;
        public IntPtr lpDirectory;
        public int nShow;
        public int dwHotKey;
        public IntPtr hIcon;
        public IntPtr lpTitle;
        public IntPtr lpVerbW;
        public IntPtr lpParametersW;
        public IntPtr lpDirectoryW;
        public IntPtr lpTitleW;
        public POINT ptInvoke;
    }

    public static int MakeHResultSuccess(ushort code) => code;

    public static bool IsVerbIdentifier(IntPtr value) => HighWord(value) == 0;

    public static ushort LowWord(IntPtr value) => (ushort)(value.ToInt64() & 0xFFFF);

    private static ushort HighWord(IntPtr value) => (ushort)((value.ToInt64() >> 16) & 0xFFFF);

    public static string? ResolveVerb(CMINVOKECOMMANDINFOEX info)
    {
        if ((info.fMask & CMIC_MASK_UNICODE) != 0 && info.lpVerbW != IntPtr.Zero)
        {
            return Marshal.PtrToStringUni(info.lpVerbW);
        }

        if (info.lpVerb != IntPtr.Zero)
        {
            return Marshal.PtrToStringAnsi(info.lpVerb);
        }

        return null;
    }
}

#endregion
