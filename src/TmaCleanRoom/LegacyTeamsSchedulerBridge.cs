using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TmaCleanRoom
{
    // Calls the scheduler through Microsoft's installed TMA assembly. Authentication
    // remains entirely inside Teams/OneAuth: this class never requests or sees a token.
    internal static class LegacyTeamsSchedulerBridge
    {
        private const string StockClsid = "{19A6E644-14E6-4A60-B8D7-DD20610A871D}";
        private const int AccountDiscoveryDelayMilliseconds = 250;
        private const int TelemetryDataCategoryRequiredServiceData = 1;
        private static string assemblyDirectory;
        private static Assembly teamsAssembly;
        private static object teamsApplication;
        private static object teamsScheduler;
        private static readonly object initializationLock = new object();
        private static readonly object logLock = new object();

        internal sealed class Result
        {
            internal string MeetingId;
            internal string JoinUrl;
            internal string BodyHtml;
            internal string BodyText;
            internal string OptionsUrl;
        }

        internal static void WarmUp()
        {
            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    var timer = Stopwatch.StartNew();
                    Log("WarmUp: start");
                    lock (initializationLock)
                    {
                        Assembly assembly = LoadTeamsAssembly();
                        Type serviceType = RequiredType(assembly,
                            "Microsoft.Teams.MeetingAddin.Service.Accounts.UserAccountService");
                        object service = InvokeStatic(serviceType, "GetInstance");
                        object[] args = { null };
                        MethodInfo tryPrimary = RequiredMethod(serviceType,
                            "TryGetPrimaryUser", 1);
                        bool found = Convert.ToBoolean(tryPrimary.Invoke(service, args));
                        if (!found || args[0] == null)
                        {
                            InitializeTeamsSettings(assembly);
                            MethodInfo loadAll = RequiredMethod(serviceType, "LoadAllUsers", 0);
                            for (int attempt = 0; attempt < 40; attempt++)
                            {
                                Thread.Sleep(AccountDiscoveryDelayMilliseconds);
                                if (attempt == 3 || attempt == 11 || attempt == 23)
                                    loadAll.Invoke(service, null);
                                args[0] = null;
                                found = Convert.ToBoolean(tryPrimary.Invoke(service, args));
                                if (found && args[0] != null) break;
                            }
                        }
                        if (found && args[0] != null)
                        {
                            EnsureOneAuthStarted(assembly);
                            object settings = GetProperty(args[0], "Settings");
                            GetCachedScheduler(assembly, service, settings);
                        }
                        Log("WarmUp: account=" + (found && args[0] != null) +
                            ", elapsedMs=" + timer.ElapsedMilliseconds);
                    }
                }
                catch (Exception exception) { LogException("WarmUp: failure", exception); }
            });
        }

        internal static Result CreateMeeting(Outlook.AppointmentItem appointment,
            string templateId)
        {
            try
            {
                Log("CreateMeeting: start");
                var timer = Stopwatch.StartNew();
                Result result = CreateMeetingCore(appointment, templateId, false);
                Log("CreateMeeting: success, elapsedMs=" + timer.ElapsedMilliseconds);
                return result;
            }
            catch (Exception exception)
            {
                LogException("CreateMeeting: failure", exception);
                throw;
            }
        }

        internal static Result CreateMeetNow(Outlook.AppointmentItem appointment)
        {
            return CreateMeetingCore(appointment, null, true);
        }

        private static Result CreateMeetingCore(Outlook.AppointmentItem appointment,
            string templateId, bool meetNow)
        {
            if (appointment == null) throw new ArgumentNullException("appointment");
            Assembly assembly = LoadTeamsAssembly();
            Log("Teams assembly loaded: " + assembly.FullName);

            Type accountServiceType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Service.Accounts.UserAccountService");
            object accountService = InvokeStatic(accountServiceType, "GetInstance");
            Log("UserAccountService acquired");
            object[] primaryArgs = { null };
            MethodInfo tryPrimary = RequiredMethod(accountServiceType,
                "TryGetPrimaryUser", 1);
            bool found = Convert.ToBoolean(tryPrimary.Invoke(accountService, primaryArgs));
            if (!found || primaryArgs[0] == null)
            {
                Log("Primary account absent; initializing Teams settings");
                InitializeTeamsSettings(assembly);
                MethodInfo loadAllUsers = RequiredMethod(accountServiceType,
                    "LoadAllUsers", 0);
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    Thread.Sleep(AccountDiscoveryDelayMilliseconds);
                    if (attempt == 3 || attempt == 11 || attempt == 23)
                    {
                        loadAllUsers.Invoke(accountService, null);
                        Log("UserAccountService.LoadAllUsers invoked at attempt " + attempt);
                    }
                    primaryArgs[0] = null;
                    found = Convert.ToBoolean(tryPrimary.Invoke(accountService, primaryArgs));
                    if (found && primaryArgs[0] != null) break;
                }
            }
            object account = primaryArgs[0];
            Log("Primary account lookup complete: found=" + found + ", object=" + (account != null));
            if (!found || account == null || !GetBoolean(account, "IsValidUser"))
                throw new InvalidOperationException(
                    "Teams ne fournit aucun compte valide au composant Meeting Add-in. " +
                    "Connectez Teams avec le compte Office puis reessayez.");
            EnsureOneAuthStarted(assembly);

            object settings = GetProperty(account, "Settings");
            string audience = Convert.ToString(GetProperty(settings, "AudienceUrl"));
            if (String.IsNullOrWhiteSpace(audience))
                throw new InvalidOperationException("Le compte Teams ne fournit pas AudienceUrl.");

            object telemetry = CreateTelemetryContext(assembly);
            Log("Telemetry context acquired: " + (telemetry != null));
            Type contextType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Scheduler.RequestContext");
            object context = Activator.CreateInstance(contextType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                null, new[] { telemetry }, null);
            SetProperty(context, "TeamsUserAccount", account);
            SetProperty(context, "ResourceUri", audience);
            SetProperty(context, "UsePost", true);
            SetProperty(context, "IsChannelMeeting", false);

            object request = BuildRequest(assembly, appointment, account,
                templateId, meetNow);
            Log("Meeting request built");
            object scheduler = GetCachedScheduler(assembly, accountService, settings);
            object taskObject = RequiredMethod(scheduler.GetType(), "CreateMeetingAsync", 2)
                .Invoke(scheduler, new[] { request, context });
            Log("Scheduler task started");
            Task task = taskObject as Task;
            if (task == null) throw new InvalidOperationException("Teams n'a pas retourne une tache scheduler.");
            task.GetAwaiter().GetResult();
            object metadata = GetProperty(taskObject, "Result");

            string join = Convert.ToString(GetProperty(metadata, "OnlineMeetingConfLink"));
            if (String.IsNullOrWhiteSpace(join))
            {
                object links = GetProperty(metadata, "Links");
                if (links != null) join = Convert.ToString(GetProperty(links, "Join"));
                if (String.IsNullOrWhiteSpace(join) && links != null)
                    join = Convert.ToString(GetProperty(links, "ShortOnlineMeetingJoinUrl"));
            }
            string id = Convert.ToString(GetProperty(metadata, "NumericMeetingId"));
            if (String.IsNullOrWhiteSpace(join))
                throw new InvalidOperationException("Le scheduler Teams n'a retourne aucun lien de reunion.");
            if (String.IsNullOrWhiteSpace(id)) id = Guid.NewGuid().ToString("N");
            object body = GetProperty(metadata, "Body");
            object meetingLinks = GetProperty(metadata, "Links");
            string optionsUrl = meetingLinks == null ? null :
                Convert.ToString(GetProperty(meetingLinks, "Options"));
            string rawHtml = body == null ? null : Convert.ToString(GetProperty(body, "Html"));
            string bodyHtml = DecodeInvitationBody(rawHtml);
            string bodyText = body == null ? null : DecodeInvitationBody(
                Convert.ToString(GetProperty(body, "Text")));
            Log("Scheduler body received: html=" + !String.IsNullOrWhiteSpace(bodyHtml) +
                ", text=" + !String.IsNullOrWhiteSpace(bodyText) +
                ", media=" + GetDataMediaType(rawHtml) +
                ", images=" + CountOccurrences(bodyHtml, "<img"));
            return new Result { MeetingId = id, JoinUrl = join,
                BodyHtml = bodyHtml, BodyText = bodyText, OptionsUrl = optionsUrl };
        }

        // Mirrors Microsoft.Teams.MeetingAddin.View.Invitation.MeetingInvitation
        // FormatBody: scheduler bodies are data URIs, not directly insertable HTML.
        private static string DecodeInvitationBody(string rawBody)
        {
            if (String.IsNullOrEmpty(rawBody)) return rawBody;
            string decoded = Uri.UnescapeDataString(rawBody.Replace('+', ' '));
            int separator = decoded.IndexOf(',');
            return separator == -1 ? decoded : decoded.Substring(separator + 1);
        }

        private static string GetDataMediaType(string rawBody)
        {
            if (String.IsNullOrEmpty(rawBody) || !rawBody.StartsWith("data:",
                StringComparison.OrdinalIgnoreCase)) return "none";
            int separator = rawBody.IndexOf(',');
            return separator < 0 ? "invalid" : rawBody.Substring(5,
                Math.Min(separator - 5, 80));
        }

        private static int CountOccurrences(string value, string marker)
        {
            if (String.IsNullOrEmpty(value)) return 0;
            int count = 0, offset = 0;
            while ((offset = value.IndexOf(marker, offset,
                StringComparison.OrdinalIgnoreCase)) >= 0)
            { count++; offset += marker.Length; }
            return count;
        }

        private static void InitializeTeamsSettings(Assembly assembly)
        {
            EnsureTeamsApplicationContext(assembly);
            Type factoryType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Settings.SettingsFilesMonitorFactory");
            object monitor = InvokeStatic(factoryType, "GetSettingsFilesMonitor");
            Log("Settings monitor acquired: " + (monitor != null));
            Type appSettingsType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.AppSettings");
            object settings = appSettingsType.GetProperty("Instance",
                BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                .GetValue(null, null);
            Log("AppSettings acquired: " + (settings != null));
            MethodInfo initialize = RequiredMethod(settings.GetType(), "Initialize", 1);
            initialize.Invoke(settings, new[] { monitor });
            Log("AppSettings.Initialize returned");
        }

        private static void EnsureTeamsApplicationContext(Assembly assembly)
        {
            if (teamsApplication != null) return;
            Type globalsType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Globals");
            FieldInfo meetingAddin = globalsType.GetField("MeetingAddin",
                BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            object existing = meetingAddin == null ? null : meetingAddin.GetValue(null);
            if (existing != null)
            {
                teamsApplication = existing;
                Log("Teams Application context already exists");
                return;
            }
            Type applicationType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Application");
            object created = Activator.CreateInstance(applicationType);
            if (GetBoolean(created, "IsApplicationInBadState"))
                throw new InvalidOperationException(
                    "Le contexte Application des DLL Teams n'a pas pu s'initialiser.");
            teamsApplication = created;
            Log("Teams Application context created without OnStartup");
        }

        private static void EnsureOneAuthStarted(Assembly assembly)
        {
            Type authenticatorType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Scheduler.OneAuthAuthenticator");
            FieldInfo startedField = authenticatorType.GetField("HasOneAuthStarted",
                BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            if (startedField != null && Convert.ToBoolean(startedField.GetValue(null)))
            {
                Log("OneAuth already started");
                return;
            }
            Type applicationType = teamsApplication.GetType();
            object lifecycle = RequiredField(applicationType, "telemetryLifeCycleContext")
                .GetValue(teamsApplication);
            object hrdService = RequiredField(applicationType, "hrdHostService")
                .GetValue(teamsApplication);
            object logger = RequiredField(applicationType, "logger").GetValue(null);
            MethodInfo startup = RequiredMethod(authenticatorType, "OneAuthStartup", 3);
            startup.Invoke(null, new[] { lifecycle, logger, hrdService });
            bool started = startedField != null && Convert.ToBoolean(startedField.GetValue(null));
            Log("OneAuthStartup returned; started=" + started);
            if (!started)
            {
                Type utilsType = RequiredType(assembly,
                    "Microsoft.Teams.MeetingAddin.Scheduler.OneAuthUtils");
                string lastError = Convert.ToString(RequiredMethod(utilsType,
                    "GetLastOneAuthError", 0).Invoke(null, null));
                Log("OneAuth startup status: " + SanitizeOneAuthStatus(lastError));
                if (!String.IsNullOrEmpty(lastError) &&
                    (lastError.IndexOf("2401", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     lastError.IndexOf("double initialize", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     String.Equals(lastError.Trim(), "6::0", StringComparison.Ordinal)))
                {
                    startedField.SetValue(null, true);
                    started = true;
                    Log("OneAuth native instance already active; managed state synchronized");
                }
            }
            if (!started)
                throw new InvalidOperationException("Le demarrage OneAuth des DLL Teams a echoue.");
        }

        private static object GetCachedScheduler(Assembly assembly,
            object accountService, object settings)
        {
            if (teamsScheduler != null) return teamsScheduler;
            lock (initializationLock)
            {
                if (teamsScheduler != null) return teamsScheduler;
                Type timerFactoryType = RequiredType(assembly,
                    "Microsoft.Teams.MeetingAddin.Timer.TimerFactory");
                object timerFactory = timerFactoryType.GetProperty("Instance",
                    BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                    .GetValue(null, null);
                Type licenseType = RequiredType(assembly,
                    "Microsoft.Teams.MeetingAddin.Service.Accounts.CopilotLicenseService");
                object license = RequiredMethod(licenseType, "Create", 0)
                    .Invoke(null, null);
                Type schedulerType = RequiredType(assembly,
                    "Microsoft.Teams.MeetingAddin.Scheduler.SchedulerServiceAsyncWithCache");
                object scheduler = Activator.CreateInstance(schedulerType,
                    new[] { accountService, timerFactory, license });
                RequiredMethod(schedulerType, "CreateCacheFolderForUser", 1)
                    .Invoke(scheduler, new[] { settings });
                teamsScheduler = scheduler;
                Log("SchedulerServiceAsyncWithCache initialized");
                return scheduler;
            }
        }

        private static string SanitizeOneAuthStatus(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return "empty";
            // Startup status contains an error code/message only. Keep it bounded so
            // unexpected native output cannot turn the diagnostic into a data dump.
            value = value.Replace('\r', ' ').Replace('\n', ' ');
            return value.Length <= 500 ? value : value.Substring(0, 500);
        }

        internal static void Log(string message)
        {
            try
            {
                // Warm-up and ribbon actions can log from different threads. Keep
                // rotation and append atomic so one writer cannot move the file while
                // another writer is opening it.
                lock (logLock)
                {
                    string directory = Path.Combine(Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData), "TMA-CleanRoom");
                    Directory.CreateDirectory(directory);
                    string path = Path.Combine(directory, "teams-bridge.log");
                    if (File.Exists(path) && new FileInfo(path).Length >= 2 * 1024 * 1024)
                    {
                        string previous = Path.Combine(directory, "teams-bridge.previous.log");
                        if (File.Exists(previous)) File.Delete(previous);
                        File.Move(path, previous);
                    }
                    File.AppendAllText(path,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " [" +
                        AppDomain.CurrentDomain.FriendlyName + "] " + message + Environment.NewLine);
                }
            }
            catch { }
        }

        internal static void LogException(string message, Exception exception)
        {
            Log(message + Environment.NewLine + exception);
            Exception inner = exception.InnerException;
            int depth = 1;
            while (inner != null)
            {
                Log("Inner exception " + depth + Environment.NewLine + inner);
                inner = inner.InnerException;
                depth++;
            }
        }

        private static object BuildRequest(Assembly assembly, Outlook.AppointmentItem appointment,
            object account, string templateId, bool meetNow)
        {
            Type requestType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Scheduler.MeetingRequest");
            object request = Activator.CreateInstance(requestType);
            SetProperty(request, "Subject", appointment.Subject ?? String.Empty);
            SetProperty(request, "StartTime", appointment.Start.ToUniversalTime());
            SetProperty(request, "EndTime", appointment.End.ToUniversalTime());
            SetProperty(request, "Location", appointment.Location ?? String.Empty);
            PropertyInfo meetingType = RequiredProperty(requestType, "MeetingType");
            meetingType.SetValue(request, Enum.ToObject(meetingType.PropertyType,
                meetNow ? 4 : 1), null);

            if (!String.IsNullOrWhiteSpace(templateId))
            {
                Type detailsType = RequiredType(assembly,
                    "Microsoft.Teams.MeetingAddin.Model.TemplateDetails");
                object details = Activator.CreateInstance(detailsType);
                SetProperty(details, "id", templateId);
                SetProperty(request, "templateDetails", details);
                Log("Meeting template selected: " + templateId);
            }

            Type participantsType = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Scheduler.MeetingParticipants");
            object participants = Activator.CreateInstance(participantsType);
            SetProperty(participants, "Organizer", Convert.ToString(GetProperty(account, "Name")));
            SetProperty(participants, "Attendees", ReadAttendees(appointment));
            SetProperty(request, "Participants", participants);
            return request;
        }

        private static string[] ReadAttendees(Outlook.AppointmentItem appointment)
        {
            var result = new List<string>();
            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
                for (int i = 1; i <= recipients.Count; i++)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = recipients[i];
                        string address = recipient.Address;
                        if (!String.IsNullOrWhiteSpace(address)) result.Add(address);
                    }
                    finally { if (recipient != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(recipient); }
                }
            }
            finally { if (recipients != null) System.Runtime.InteropServices.Marshal.FinalReleaseComObject(recipients); }
            return result.ToArray();
        }

        private static object CreateTelemetryContext(Assembly assembly)
        {
            Type manager = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Telemetry.TelemetryManager");
            PropertyInfo current = manager.GetProperty("CurrentContext",
                BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            object context = current == null ? null : current.GetValue(null, null);
            if (context != null) return context;
            Type category = RequiredType(assembly,
                "Microsoft.Teams.MeetingAddin.Telemetry.DataCategory");
            MethodInfo create = RequiredMethod(manager, "CreateContext", 2);
            return create.Invoke(null, new[] { "TmaCleanRoom.CreateMeeting",
                Enum.ToObject(category, TelemetryDataCategoryRequiredServiceData) });
        }

        private static Assembly LoadTeamsAssembly()
        {
            if (teamsAssembly != null) return teamsAssembly;
            lock (initializationLock)
            {
                if (teamsAssembly != null) return teamsAssembly;
                string localDirectory = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location);
                string localMeetingAddin = Path.Combine(localDirectory,
                    "Microsoft.Teams.MeetingAddin.dll");
                if (File.Exists(localMeetingAddin))
                {
                    assemblyDirectory = localDirectory;
                    Log("Using standalone Teams payload: " + assemblyDirectory);
                }
                else
                {
                    string loader = ReadStockLoaderPath();
                    assemblyDirectory = Path.GetDirectoryName(loader);
                    Log("Using installed Teams payload fallback: " + assemblyDirectory);
                }
                string path = Path.Combine(assemblyDirectory,
                    "Microsoft.Teams.MeetingAddin.dll");
                if (!File.Exists(path))
                    throw new FileNotFoundException(
                        "DLL Meeting Add-in Teams introuvable.", path);

                // Register one resolver for the lifetime of the AppDomain. The old
                // implementation added a new anonymous handler on every meeting.
                AppDomain.CurrentDomain.AssemblyResolve += ResolveTeamsDependency;
                teamsAssembly = Assembly.LoadFrom(path);
                return teamsAssembly;
            }
        }

        private static Assembly ResolveTeamsDependency(object sender,
            ResolveEventArgs args)
        {
            string simpleName = new AssemblyName(args.Name).Name;
            string dependency = Path.Combine(assemblyDirectory, simpleName + ".dll");
            return File.Exists(dependency) ? Assembly.LoadFrom(dependency) : null;
        }

        private static string ReadStockLoaderPath()
        {
            string keyPath = @"Software\Classes\CLSID\" + StockClsid + @"\InprocServer32";
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                string value = key == null ? null : Convert.ToString(key.GetValue(null));
                if (!String.IsNullOrWhiteSpace(value) && File.Exists(value)) return value;
            }
            throw new InvalidOperationException("L'enregistrement COM du Meeting Add-in Teams stock est introuvable.");
        }

        private static Type RequiredType(Assembly assembly, string name)
        { return assembly.GetType(name, true, false); }
        private static MethodInfo RequiredMethod(Type type, string name, int parameterCount)
        {
            foreach (MethodInfo method in type.GetMethods(BindingFlags.Instance | BindingFlags.Static |
                BindingFlags.Public | BindingFlags.NonPublic))
                if (method.Name == name && method.GetParameters().Length == parameterCount) return method;
            throw new MissingMethodException(type.FullName, name);
        }
        private static PropertyInfo RequiredProperty(Type type, string name)
        {
            PropertyInfo property = type.GetProperty(name,
                BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            if (property == null) throw new MissingMemberException(type.FullName, name);
            return property;
        }
        private static FieldInfo RequiredField(Type type, string name)
        {
            FieldInfo field = type.GetField(name,
                BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            if (field == null) throw new MissingFieldException(type.FullName, name);
            return field;
        }
        private static object GetProperty(object target, string name)
        { return target == null ? null : RequiredProperty(target.GetType(), name).GetValue(target, null); }
        private static bool GetBoolean(object target, string name)
        { return Convert.ToBoolean(GetProperty(target, name)); }
        private static void SetProperty(object target, string name, object value)
        { RequiredProperty(target.GetType(), name).SetValue(target, value, null); }
        private static object InvokeStatic(Type type, string name)
        { return RequiredMethod(type, name, 0).Invoke(null, null); }
    }
}
