using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace TmaCleanRoom
{
    internal static class OfficeNativeSignIn
    {
        private const int ShowUiOrdinal = 44059;
        private const int ShowUiParamsSize = 0xA0;
        private const int OsfCurrentIdentityRva = 0x288190;
        private const string ValidatedOsfVersion = "16.0.20228.20014";

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
