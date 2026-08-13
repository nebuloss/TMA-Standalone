using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace TmaCleanRoom
{
    // Thin wrapper around Office's already loaded native identity components.
    // It intentionally does not acquire OAuth tokens or register an application.
    // Native entry points are validated before invocation because these functions
    // are undocumented and their ordinals/layouts may change after an Office update.
    internal static class OfficeNativeSignIn
    {
        private const string StandaloneRegistryPath = @"Software\TMA-Standalone";
        private const string SignedOutValue = "AddinSignedOut";

        private const int DosHeaderPePointerOffset = 0x3c;
        private const int PeSignature = 0x00004550;
        private const short Pe32Magic = 0x10b;
        private const short Pe32PlusMagic = 0x20b;
        private const int Pe32DataDirectoryOffset = 96;
        private const int Pe32PlusDataDirectoryOffset = 112;
        private const int PeFileHeaderSize = 24;
        private const int ExportNumberOfFunctionsOffset = 20;
        private const int ExportAddressOfFunctionsOffset = 28;
        private const int MaximumReasonableExportCount = 100000;
        private static readonly byte[] ShowUiX64Signature =
            { 0x4c, 0x8b, 0xc9, 0x4c, 0x8b, 0xc2 };

        private enum OfficeSignInMode : long
        {
            // Value constructed by Office for its interactive account picker.
            InteractiveAccountPicker = 7
        }

        [StructLayout(LayoutKind.Explicit, Size = 160)]
        private struct ShowUiParameters
        {
            // These fields were identified from the native ShowUIParams constructor.
            // All unlisted bytes are reserved and remain zero-initialized.
            [FieldOffset(40)]
            internal OfficeSignInMode PrimaryMode;

            [FieldOffset(88)]
            internal OfficeSignInMode SecondaryMode;

            [FieldOffset(116)]
            [MarshalAs(UnmanagedType.I1)]
            internal bool AllowInteractiveUi;

            internal static ShowUiParameters CreateInteractive()
            {
                return new ShowUiParameters
                {
                    PrimaryMode = OfficeSignInMode.InteractiveAccountPicker,
                    SecondaryMode = OfficeSignInMode.InteractiveAccountPicker,
                    AllowInteractiveUi = true
                };
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int ShowUiDelegate(IntPtr parameters, out IntPtr identity);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);

        internal static string Show()
        {
            IntPtr module = GetModuleHandle("mso98win32client.dll");
            if (module == IntPtr.Zero)
                throw new InvalidOperationException("Le composant d'identite Office n'est pas charge dans Outlook.");

            ProcessModule loadedModule = FindLoadedModule(module);
            string version = loadedModule.FileVersionInfo.FileVersion;

            IntPtr address = FindUniqueShowUiExport(loadedModule);

            IntPtr parameters = Marshal.AllocHGlobal(
                Marshal.SizeOf(typeof(ShowUiParameters)));
            IntPtr identity = IntPtr.Zero;
            try
            {
                ShowUiParameters nativeParameters =
                    ShowUiParameters.CreateInteractive();
                Marshal.StructureToPtr(nativeParameters, parameters, false);

                ShowUiDelegate showUi = (ShowUiDelegate)Marshal.GetDelegateForFunctionPointer(
                    address, typeof(ShowUiDelegate));
                int result = showUi(parameters, out identity);
                if (result < 0)
                    Marshal.ThrowExceptionForHR(result);
                // A successful ShowUI call may return before Outlook refreshes its
                // Accounts collection. Connection state is therefore evaluated by
                // the documented Outlook Object Model after this method returns.
                return "Fenetre de connexion Office ouverte.\r\n\r\nModule MSO : " +
                    version;
            }
            finally
            {
                if (identity != IntPtr.Zero) Marshal.Release(identity);
                Marshal.FreeHGlobal(parameters);
            }
        }

        internal static bool IsStandaloneEnabled()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(
                StandaloneRegistryPath))
            {
                return key == null || Convert.ToInt32(
                    key.GetValue(SignedOutValue, 0)) == 0;
            }
        }

        internal static void SetStandaloneEnabled(bool enabled)
        {
            // This flag signs out only the standalone add-in. It deliberately does
            // not alter Office identities, Outlook profiles, or cached credentials.
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(
                StandaloneRegistryPath))
            {
                key.SetValue(SignedOutValue, enabled ? 0 : 1,
                    RegistryValueKind.DWord);
            }
        }

        internal static string GetProfilePicturePath()
        {
            string picturesDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft", "Office", "16.0", "People", "Pictures");
            if (!Directory.Exists(picturesDirectory)) return null;

            // Office stores the signed-in user's downloaded persona image here. A
            // profile can briefly leave an older file behind while refreshing, so
            // prefer the most recently updated supported image.
            return Directory.EnumerateFiles(picturesDirectory)
                .Where(path => String.Equals(Path.GetExtension(path), ".jpg",
                                   StringComparison.OrdinalIgnoreCase) ||
                               String.Equals(Path.GetExtension(path), ".jpeg",
                                   StringComparison.OrdinalIgnoreCase) ||
                               String.Equals(Path.GetExtension(path), ".png",
                                   StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();
        }

        private static ProcessModule FindLoadedModule(IntPtr baseAddress)
        {
            foreach (ProcessModule candidate in Process.GetCurrentProcess().Modules)
            {
                if (candidate.BaseAddress == baseAddress) return candidate;
            }
            throw new InvalidOperationException("Impossible de determiner le chemin du composant Office.");
        }

        private static IntPtr FindUniqueShowUiExport(ProcessModule module)
        {
            IntPtr imageBase = module.BaseAddress;
            int imageSize = module.ModuleMemorySize;
            int peOffset = ReadInt32(imageBase, imageSize,
                DosHeaderPePointerOffset);
            if (ReadInt32(imageBase, imageSize, peOffset) != PeSignature)
                throw new BadImageFormatException(
                    "Le module d'identite Office n'est pas une image PE valide.");

            int optionalHeader = peOffset + PeFileHeaderSize;
            short magic = ReadInt16(imageBase, imageSize, optionalHeader);
            int dataDirectory = optionalHeader + (magic == Pe32PlusMagic ?
                Pe32PlusDataDirectoryOffset : magic == Pe32Magic ?
                Pe32DataDirectoryOffset : throw new BadImageFormatException(
                    "Format PE Office non pris en charge."));
            int exportRva = ReadInt32(imageBase, imageSize, dataDirectory);
            if (exportRva == 0)
                throw new MissingMethodException(
                    "Le module d'identite Office n'exporte aucune fonction.");

            int functionCount = ReadInt32(imageBase, imageSize,
                exportRva + ExportNumberOfFunctionsOffset);
            int functionsRva = ReadInt32(imageBase, imageSize,
                exportRva + ExportAddressOfFunctionsOffset);
            if (functionCount <= 0 || functionCount > MaximumReasonableExportCount)
                throw new BadImageFormatException(
                    "Table d'exports Office incoherente.");

            IntPtr match = IntPtr.Zero;
            for (int index = 0; index < functionCount; index++)
            {
                int functionRva = ReadInt32(imageBase, imageSize,
                    functionsRva + index * sizeof(int));
                if (functionRva <= 0 ||
                    functionRva + ShowUiX64Signature.Length > imageSize)
                    continue;
                bool matches = true;
                for (int offset = 0; offset < ShowUiX64Signature.Length; offset++)
                {
                    if (Marshal.ReadByte(imageBase, functionRva + offset) !=
                        ShowUiX64Signature[offset])
                    { matches = false; break; }
                }
                if (!matches) continue;
                if (match != IntPtr.Zero)
                    throw new NotSupportedException(
                        "Plusieurs fonctions Office correspondent a ShowUI; appel refuse.");
                match = IntPtr.Add(imageBase, functionRva);
            }
            if (match == IntPtr.Zero)
                throw new NotSupportedException(
                    "Aucune fonction Office compatible avec SignIn::ShowUI n'a ete trouvee.");
            LegacyTeamsSchedulerBridge.Log(
                "Office SignIn::ShowUI resolved from validated PE export signature");
            return match;
        }

        private static int ReadInt32(IntPtr imageBase, int imageSize, int offset)
        {
            if (offset < 0 || offset > imageSize - sizeof(int))
                throw new BadImageFormatException("Lecture hors limites du module Office.");
            return Marshal.ReadInt32(imageBase, offset);
        }

        private static short ReadInt16(IntPtr imageBase, int imageSize, int offset)
        {
            if (offset < 0 || offset > imageSize - sizeof(short))
                throw new BadImageFormatException("Lecture hors limites du module Office.");
            return Marshal.ReadInt16(imageBase, offset);
        }
    }
}
