using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using IBindCtx = System.Runtime.InteropServices.ComTypes.IBindCtx;

namespace CropperShellExtension;

[ComVisible(true)]
[Guid(ClassGuid)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class CropperContextMenuExtension : IExplorerCommand
{
    private const string ClassGuid = "7F56439E-2130-4115-A27A-8D562049B848";
    private const string MenuText = "Crop with PyCropper";
    private const string HelpText = "Crop selected file with PyCropper";
    private const string CropperExecutable = "cropper.exe";
    private const string ExplorerCommandKey = @"*\shell\PyCropper";
    private const int DefaultTolerance = 100;

    private static readonly int[] Tolerances = { 50, 100, 150, 200 };

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
        LaunchCropper(ExtractFilePaths(psiItemArray), DefaultTolerance);
    }

    void IExplorerCommand.GetFlags(out EXPCMDFLAGS pFlags)
    {
        pFlags = EXPCMDFLAGS.ECF_HASSUBCOMMANDS | EXPCMDFLAGS.ECF_HASSUBMENU | EXPCMDFLAGS.ECF_ALWAYS_SHOW;
    }

    void IExplorerCommand.EnumSubCommands(out IEnumExplorerCommand? ppEnum)
    {
        var commands = new IExplorerCommand[Tolerances.Length];
        for (var i = 0; i < Tolerances.Length; i++)
        {
            commands[i] = new ToleranceCommand(Tolerances[i]);
        }

        ppEnum = new ExplorerCommandEnumerator(commands);
    }

    [ComRegisterFunction]
    public static void Register(Type _)
    {
        using var explorerCommandKey = Registry.ClassesRoot.CreateSubKey(ExplorerCommandKey);
        if (explorerCommandKey is null)
        {
            return;
        }

        explorerCommandKey.SetValue("MUIVerb", MenuText);
        explorerCommandKey.SetValue("ExplorerCommandHandler", $"{{{ClassGuid}}}");
    }

    [ComUnregisterFunction]
    public static void Unregister(Type _)
    {
        try
        {
            Registry.ClassesRoot.DeleteSubKeyTree(ExplorerCommandKey, throwOnMissingSubKey: false);
        }
        catch (ArgumentException)
        {
        }
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

    private static void LaunchCropper(IEnumerable<string> selection, int tolerance)
    {
        var target = GetFirstPath(selection);
        if (target is null)
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

            var toleranceArgument = tolerance.ToString(CultureInfo.InvariantCulture);
            info.ArgumentList.Add("-t");
            info.ArgumentList.Add(toleranceArgument);
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

    private static string? GetFirstPath(IEnumerable<string> selection)
    {
        foreach (var candidate in selection)
        {
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static IntPtr AllocateCoTaskMem(string text)
    {
        return Marshal.StringToCoTaskMemUni(text ?? string.Empty);
    }

    private sealed class ToleranceCommand : IExplorerCommand
    {
        private readonly int _tolerance;
        private readonly string _title;

        internal ToleranceCommand(int tolerance)
        {
            _tolerance = tolerance;
            _title = $"Tolerance {_tolerance}";
        }

        public void GetTitle(IShellItemArray? psiItemArray, out IntPtr ppszName)
        {
            ppszName = AllocateCoTaskMem(_title);
        }

        public void GetIcon(IShellItemArray? psiItemArray, out IntPtr ppszIcon)
        {
            ppszIcon = IntPtr.Zero;
        }

        public void GetToolTip(IShellItemArray? psiItemArray, out IntPtr ppszInfotip)
        {
            ppszInfotip = AllocateCoTaskMem($"{HelpText} ({_title})");
        }

        public void GetCanonicalName(out Guid pguidCommandName)
        {
            pguidCommandName = Guid.Empty;
        }

        public void GetState(IShellItemArray? psiItemArray, bool fOkToBeSlow, out EXPCMDSTATE pCmdState)
        {
            pCmdState = HasSelection(psiItemArray) ? EXPCMDSTATE.ECS_ENABLED : EXPCMDSTATE.ECS_HIDDEN;
        }

        public void Invoke(IShellItemArray? psiItemArray, IBindCtx? pbc)
        {
            LaunchCropper(ExtractFilePaths(psiItemArray), _tolerance);
        }

        public void GetFlags(out EXPCMDFLAGS pFlags)
        {
            pFlags = EXPCMDFLAGS.ECF_DEFAULT;
        }

        public void EnumSubCommands(out IEnumExplorerCommand? ppEnum)
        {
            ppEnum = null;
        }
    }

    private sealed class ExplorerCommandEnumerator : IEnumExplorerCommand
    {
        private readonly IExplorerCommand[] _commands;
        private int _index;

        internal ExplorerCommandEnumerator(IExplorerCommand[] commands)
        {
            _commands = commands;
        }

        public void Next(uint celt, out IExplorerCommand? pUICommand, out uint pceltFetched)
        {
            if (_index >= _commands.Length)
            {
                pUICommand = null;
                pceltFetched = 0;
                return;
            }

            pUICommand = _commands[_index++];
            pceltFetched = 1;
        }

        public void Skip(uint celt)
        {
            var nextIndex = _index + (int)celt;
            _index = nextIndex > _commands.Length ? _commands.Length : nextIndex;
        }

        public void Reset()
        {
            _index = 0;
        }

        public void Clone(out IEnumExplorerCommand? ppenum)
        {
            var clone = new ExplorerCommandEnumerator(_commands)
            {
                _index = _index
            };
            ppenum = clone;
        }
    }
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
