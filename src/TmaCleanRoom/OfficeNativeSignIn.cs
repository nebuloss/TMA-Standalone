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
        private const int ShowUiOrdinal = 44059;
        private const int ShowUiParamsSize = 0xA0;
        private const int OsfCurrentIdentityRva = 0x288190;
        private const string ValidatedOsfVersion = "16.0.20228.20014";
        private const string StandaloneRegistryPath = @"Software\TMA-Standalone";
        private const string SignedOutValue = "AddinSignedOut";

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int ShowUiDelegate(IntPtr parameters, out IntPtr identity);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate IntPtr GetCurrentIdentityDelegate();

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr module, IntPtr ordinal);

        internal static string Show()
        {
            IntPtr module = GetModuleHandle("mso98win32client.dll");
            if (module == IntPtr.Zero)
                throw new InvalidOperationException("Le composant d'identite Office n'est pas charge dans Outlook.");

            ProcessModule loadedModule = FindLoadedModule(module);
            string version = loadedModule.FileVersionInfo.FileVersion;

            IntPtr address = GetProcAddress(module, new IntPtr(ShowUiOrdinal));
            if (address == IntPtr.Zero)
                throw new InvalidOperationException("Mso::SignIn::ShowUI (ordinal 44059) est introuvable.");
            ValidateShowUiPrologue(address);

            IntPtr parameters = Marshal.AllocHGlobal(ShowUiParamsSize);
            IntPtr identity = IntPtr.Zero;
            bool usedProfileFallback = false;
            try
            {
                for (int offset = 0; offset < ShowUiParamsSize; offset += sizeof(long))
                    Marshal.WriteInt64(parameters, offset, 0);

                // Layout produit par ShowUIParams(HWND, ULONG) dans Office 16.0.20228.20014.
                Marshal.WriteInt64(parameters, 0x28, 7);
                Marshal.WriteInt64(parameters, 0x58, 7);
                Marshal.WriteByte(parameters, 0x74, 1);

                ShowUiDelegate showUi = (ShowUiDelegate)Marshal.GetDelegateForFunctionPointer(
                    address, typeof(ShowUiDelegate));
                int result = showUi(parameters, out identity);
                if (result < 0)
                    Marshal.ThrowExceptionForHR(result);
                if (identity == IntPtr.Zero)
                {
                    identity = GetCurrentOfficeIdentity();
                    usedProfileFallback = identity != IntPtr.Zero;
                }
                if (identity == IntPtr.Zero)
                    throw new InvalidOperationException("Office n'a retourne aucune identite (module " +
                        version + "). Aucun compte Office actif n'est disponible dans le profil Outlook.");
                return "Identite Office disponible.\r\n\r\nSource : " +
                    (usedProfileFallback ? "profil Office actif" : "fenetre de connexion Office") +
                    "\r\nModule MSO : " + version +
                    "\r\nModule OSF : " + GetLoadedModuleVersion("osf99.dll");
            }
            finally
            {
                if (identity != IntPtr.Zero) Marshal.Release(identity);
                Marshal.FreeHGlobal(parameters);
            }
        }

        internal static bool IsConnected()
        {
            IntPtr identity = IntPtr.Zero;
            try
            {
                identity = GetCurrentOfficeIdentity();
                return identity != IntPtr.Zero;
            }
            finally
            {
                if (identity != IntPtr.Zero) Marshal.Release(identity);
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

        internal static string GetAccountLabel()
        {
            const string identitiesPath = @"Software\Microsoft\Office\16.0\Common\Identity\Identities";
            using (RegistryKey identities = Registry.CurrentUser.OpenSubKey(identitiesPath))
            {
                if (identities == null) return "Compte Office actif";
                string selected = null;
                foreach (string name in identities.GetSubKeyNames())
                {
                    using (RegistryKey identity = identities.OpenSubKey(name))
                    {
                        string email = identity == null ? null : identity.GetValue("EmailAddress") as string;
                        if (String.IsNullOrWhiteSpace(email)) continue;
                        if (selected != null && !String.Equals(selected, email,
                            StringComparison.OrdinalIgnoreCase)) return "Plusieurs comptes Office actifs";
                        selected = email;
                    }
                }
                return String.IsNullOrWhiteSpace(selected) ? "Compte Office actif" : selected;
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

        private static string GetLoadedModuleVersion(string moduleName)
        {
            IntPtr module = GetModuleHandle(moduleName);
            return module == IntPtr.Zero ? "non charge" :
                FindLoadedModule(module).FileVersionInfo.FileVersion;
        }

        private static IntPtr GetCurrentOfficeIdentity()
        {
            IntPtr module = GetModuleHandle("osf99.dll");
            if (module == IntPtr.Zero) return IntPtr.Zero;

            ProcessModule loadedModule = FindLoadedModule(module);
            string version = loadedModule.FileVersionInfo.FileVersion;
            // Unlike the exported ShowUI function, this function is addressed by an
            // RVA. Calling it on an unvalidated build could execute unrelated code.
            if (!String.Equals(version, ValidatedOsfVersion, StringComparison.OrdinalIgnoreCase))
                return IntPtr.Zero;

            IntPtr address = IntPtr.Add(module, OsfCurrentIdentityRva);
            byte[] expected = { 0x40, 0x53, 0x48, 0x83, 0xec, 0x20, 0x33, 0xdb };
            for (int i = 0; i < expected.Length; i++)
            {
                if (Marshal.ReadByte(address, i) != expected[i]) return IntPtr.Zero;
            }

            GetCurrentIdentityDelegate getCurrent =
                (GetCurrentIdentityDelegate)Marshal.GetDelegateForFunctionPointer(
                    address, typeof(GetCurrentIdentityDelegate));
            return getCurrent();
        }

        private static ProcessModule FindLoadedModule(IntPtr baseAddress)
        {
            foreach (ProcessModule candidate in Process.GetCurrentProcess().Modules)
            {
                if (candidate.BaseAddress == baseAddress) return candidate;
            }
            throw new InvalidOperationException("Impossible de determiner le chemin du composant Office.");
        }

        private static void ValidateShowUiPrologue(IntPtr address)
        {
            byte[] expected = { 0x4c, 0x8b, 0xc9, 0x4c, 0x8b, 0xc2 };
            for (int i = 0; i < expected.Length; i++)
            {
                if (Marshal.ReadByte(address, i) != expected[i])
                    throw new NotSupportedException("La signature native de Mso::SignIn::ShowUI a change.");
            }
        }
    }
}
