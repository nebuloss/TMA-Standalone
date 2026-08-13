using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Web;
using System.Threading.Tasks;
using Microsoft.Win32;
using Microsoft.Office.Core;
using Extensibility;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TmaCleanRoom
{
    [ComVisible(true)]
    [Guid("8F5373B8-4973-4E58-A69E-CB57AA22691C")]
    [ProgId("TmaCleanRoom.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Addin : IDTExtensibility2, IRibbonExtensibility
    {
        private const string DialogTitle = "TMA autonome";
        private const int RibbonRepaintDelayMilliseconds = 50;
        private Outlook.Application outlook;
        private readonly List<IRibbonUI> ribbons = new List<IRibbonUI>();
        private IRibbonUI explorerRibbon;
        private Outlook.Inspectors inspectors;
        private Outlook.Explorer activeExplorer;
        private readonly List<AppointmentInspectorTracker> appointmentInspectors =
            new List<AppointmentInspectorTracker>();
        // Outlook creates the Inspector asynchronously. Keep the Explorer controls
        // disabled between the ribbon click and the corresponding NewInspector event.
        private bool explorerMeetingActionActive;

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            outlook = application as Outlook.Application;
            if (outlook != null)
            {
                inspectors = outlook.Inspectors;
                inspectors.NewInspector += Inspectors_NewInspector;
                activeExplorer = outlook.ActiveExplorer();
                if (activeExplorer != null)
                    ((Outlook.ExplorerEvents_10_Event)activeExplorer).Activate += Explorer_Activate;
                for (int index = 1; index <= inspectors.Count; index++)
                    TrackAppointmentInspector(inspectors[index]);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (inspectors != null)
                inspectors.NewInspector -= Inspectors_NewInspector;
            if (activeExplorer != null)
            {
                try { ((Outlook.ExplorerEvents_10_Event)activeExplorer).Activate -= Explorer_Activate; }
                catch (COMException) { }
                catch (InvalidComObjectException) { }
                ReleaseComObject(activeExplorer);
            }
            activeExplorer = null;
            foreach (AppointmentInspectorTracker tracker in appointmentInspectors)
                tracker.Dispose();
            appointmentInspectors.Clear();
            ReleaseComObject(inspectors);
            inspectors = null;
            ReleaseComObject(outlook);
            outlook = null;
            explorerRibbon = null;
            ribbons.Clear();
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom)
        {
            LegacyTeamsSchedulerBridge.WarmUp();
        }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string ribbonId)
        {
            if (String.Equals(ribbonId, "Microsoft.Outlook.Explorer",
                StringComparison.OrdinalIgnoreCase))
            {
                return ExplorerRibbon;
            }
            if (String.Equals(ribbonId, "Microsoft.Outlook.Appointment",
                StringComparison.OrdinalIgnoreCase))
            {
                return AppointmentRibbon;
            }
            return String.Empty;
        }

        public void ExplorerRibbon_Load(IRibbonUI ui)
        {
            LegacyTeamsSchedulerBridge.Log("Explorer ribbon loaded");
            explorerRibbon = ui;
            Ribbon_Load(ui);
        }

        public void AppointmentRibbon_Load(IRibbonUI ui)
        {
            LegacyTeamsSchedulerBridge.Log("Appointment ribbon loaded");
            Ribbon_Load(ui);
        }

        private void Ribbon_Load(IRibbonUI ui)
        {
            if (ui != null) ribbons.Add(ui);
        }

        private void InvalidateRibbons()
        {
            for (int index = ribbons.Count - 1; index >= 0; index--)
            {
                try { ribbons[index].Invalidate(); }
                catch (COMException) { ribbons.RemoveAt(index); }
                catch (InvalidComObjectException) { ribbons.RemoveAt(index); }
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            TrackAppointmentInspector(inspector);
            LegacyTeamsSchedulerBridge.Log("New inspector tracked; count=" + appointmentInspectors.Count);
        }

        private void Explorer_Activate()
        {
            // Explorer.Activate is raised when focus returns from an Inspector. It is
            // a native Outlook event and avoids polling detached COM objects.
            LegacyTeamsSchedulerBridge.Log(
                "Explorer.Activate; inspectors=" + appointmentInspectors.Count);
            InvalidateExplorerRibbon();
        }

        private void AppointmentInspector_Activate()
        {
            LegacyTeamsSchedulerBridge.Log(
                "Inspector.Activate; inspectors=" + appointmentInspectors.Count);
            InvalidateExplorerRibbon();
        }

        private void TrackAppointmentInspector(Outlook.Inspector inspector)
        {
            if (inspector == null) return;
            object item = inspector.CurrentItem;
            Outlook.AppointmentItem appointment = item as Outlook.AppointmentItem;
            if (appointment == null) return;
            appointmentInspectors.Add(new AppointmentInspectorTracker(
                this, inspector, appointment));
            // NewInspector now owns the state; the temporary click guard is no longer
            // needed once Outlook has exposed the actual appointment window.
            explorerMeetingActionActive = false;
        }

        private void AppointmentInspector_Close(AppointmentInspectorTracker tracker)
        {
            appointmentInspectors.Remove(tracker);
            LegacyTeamsSchedulerBridge.Log("Inspector close received; count=" + appointmentInspectors.Count);
            tracker.Dispose();
            InvalidateRibbons();
        }

        private void InvalidateExplorerRibbon()
        {
            try
            {
                if (explorerRibbon != null)
                {
                    explorerRibbon.InvalidateControl("TmaCleanRoom.MeetNow");
                    explorerRibbon.InvalidateControl("TmaCleanRoom.MeetingSplit");
                    explorerRibbon.Invalidate();
                }
            }
            catch (COMException) { explorerRibbon = null; }
            catch (InvalidComObjectException) { explorerRibbon = null; }
        }

        public bool ExplorerMeeting_GetEnabled(IRibbonControl control)
        {
            bool enabled = !explorerMeetingActionActive &&
                appointmentInspectors.Count == 0;
            LegacyTeamsSchedulerBridge.Log("Explorer getEnabled: control=" +
                (control == null ? "null" : control.Id) + ", enabled=" + enabled +
                ", inspectors=" + appointmentInspectors.Count);
            return enabled;
        }

        private sealed class AppointmentInspectorTracker : IDisposable
        {
            private Addin owner;
            private Outlook.Inspector inspector;
            private Outlook.AppointmentItem appointment;
            private bool closing;

            internal AppointmentInspectorTracker(Addin owner, Outlook.Inspector inspector,
                Outlook.AppointmentItem appointment)
            {
                this.owner = owner;
                this.inspector = inspector;
                this.appointment = appointment;
                ((Outlook.InspectorEvents_10_Event)inspector).Close += Inspector_Close;
                ((Outlook.InspectorEvents_10_Event)inspector).Activate += Inspector_Activate;
                ((Outlook.ItemEvents_10_Event)appointment).Close += Appointment_Close;
            }

            private void Inspector_Close()
            {
                NotifyClosed("Inspector.Close");
            }

            private void Inspector_Activate()
            {
                if (owner != null) owner.AppointmentInspector_Activate();
            }

            private void Appointment_Close(ref bool cancel)
            {
                // AppointmentItem.Close is raised reliably for "Don't Save", whereas
                // Inspector.Close may arrive later (or after its RCW was disconnected).
                if (!cancel) NotifyClosed("AppointmentItem.Close");
            }

            private void NotifyClosed(string source)
            {
                // Outlook can raise both item and inspector close events for the same
                // window. Ensure cleanup and ribbon invalidation happen only once.
                if (closing) return;
                closing = true;
                LegacyTeamsSchedulerBridge.Log("Tracker close source=" + source);
                if (owner != null) owner.AppointmentInspector_Close(this);
            }

            public void Dispose()
            {
                Outlook.Inspector current = inspector;
                Outlook.AppointmentItem currentAppointment = appointment;
                inspector = null;
                appointment = null;
                owner = null;
                // Unsubscribe before releasing each RCW. Accessing CurrentItem while
                // scanning Inspectors after closure can throw InvalidComObjectException.
                if (currentAppointment != null)
                {
                    try { ((Outlook.ItemEvents_10_Event)currentAppointment).Close -= Appointment_Close; }
                    catch (COMException) { }
                    catch (InvalidComObjectException) { }
                    try { if (Marshal.IsComObject(currentAppointment)) Marshal.FinalReleaseComObject(currentAppointment); }
                    catch (InvalidComObjectException) { }
                }
                if (current != null)
                {
                    try { ((Outlook.InspectorEvents_10_Event)current).Close -= Inspector_Close; }
                    catch (COMException) { }
                    catch (InvalidComObjectException) { }
                    try { ((Outlook.InspectorEvents_10_Event)current).Activate -= Inspector_Activate; }
                    catch (COMException) { }
                    catch (InvalidComObjectException) { }
                    try { if (Marshal.IsComObject(current)) Marshal.FinalReleaseComObject(current); }
                    catch (InvalidComObjectException) { }
                }
            }
        }

        public Bitmap Ribbon_GetImage(IRibbonControl control)
        {
            if (control != null && control.Id.IndexOf("LanguageFr",
                StringComparison.OrdinalIgnoreCase) >= 0) return DrawFrenchFlag();
            if (control != null && control.Id.IndexOf("LanguageEn",
                StringComparison.OrdinalIgnoreCase) >= 0) return DrawUsFlag();
            if (control != null && control.Id.IndexOf("LanguageMenu",
                StringComparison.OrdinalIgnoreCase) >= 0)
            {
                string language = CleanRoomMeetingService.GetInvitationLanguage(
                    ResolveAppointment(control));
                return String.Equals(language, "en", StringComparison.OrdinalIgnoreCase)
                    ? DrawUsFlag() : DrawFrenchFlag();
            }
            if (control != null && control.Id.IndexOf("Account",
                StringComparison.OrdinalIgnoreCase) >= 0)
                return DrawAccountIcon();
            string directory = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);
            string asset = control != null && control.Id.IndexOf("MeetNow",
                StringComparison.OrdinalIgnoreCase) >= 0
                ? "MeetNow_Large_96.png" : "NewMeeting_Large_96.png";
            string path = Path.Combine(directory, "Assets", asset);
            if (!File.Exists(path)) return null;
            using (var source = new Bitmap(path)) return new Bitmap(source);
        }

        public void CreateMeeting(IRibbonControl control)
        { CreateMeetingWithTemplate(control, null, null); }

        public void CreateFrenchMeeting(IRibbonControl control)
        { CreateMeetingWithTemplate(control, null, "fr"); }

        public void CreateEnglishMeeting(IRibbonControl control)
        { CreateMeetingWithTemplate(control, null, "en"); }

        public void SelectFrenchLanguage(IRibbonControl control)
        { SelectInvitationLanguage(control, "fr"); }

        public void SelectEnglishLanguage(IRibbonControl control)
        { SelectInvitationLanguage(control, "en"); }

        private void SelectInvitationLanguage(IRibbonControl control, string language)
        {
            CleanRoomMeetingService.SetInvitationLanguage(
                ResolveAppointment(control), language);
            InvalidateRibbons();
        }

        public void CreateWebinar(IRibbonControl control)
        { OpenTeamsCreationFlow(control, "/l/virtualevent/new", "eventType=Webinar"); }

        public void CreateTownHall(IRibbonControl control)
        { OpenTeamsCreationFlow(control, "/l/virtualevent/new", "eventType=Townhall"); }

        public void CreateVirtualAppointment(IRibbonControl control)
        {
            OpenTeamsCreationFlow(control, "/l/meeting/new",
                "templateId=firstparty_e514e598-fba6-4e1f-b8b3-138dd3bca748");
        }

        private void OpenTeamsCreationFlow(IRibbonControl control, string path,
            string defaultQuery)
        {
            try
            {
                NameValueCollection query = HttpUtility.ParseQueryString(defaultQuery);
                query["source"] = "OutlookTMA";
                query["sourceVersion"] = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                query["correlationId"] = Guid.NewGuid().ToString();

                Outlook.AppointmentItem appointment = ResolveAppointment(control);
                if (appointment != null)
                {
                    query["startTime"] = appointment.Start.ToString("O");
                    query["endTime"] = appointment.End.ToString("O");
                }

                bool teamsProtocolRegistered;
                using (RegistryKey protocol = Registry.ClassesRoot.OpenSubKey("msteams"))
                    teamsProtocolRegistered = protocol != null;
                string url = teamsProtocolRegistered
                    ? "msteams:" + path + "?" + query
                    : "https://teams.microsoft.com" + path + "?" + query;
                LegacyTeamsSchedulerBridge.Log("Opening official Teams creation flow: " +
                    path);
                Process.Start(url);
            }
            catch (Exception exception)
            {
                LegacyTeamsSchedulerBridge.LogException(
                    "Teams creation flow failed", exception);
                MessageBox.Show(FormatException(exception), "TMA Clean Room",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void CreateMeetingWithTemplate(IRibbonControl control, string templateId,
            string invitationLanguage)
        {
            try
            {
                Outlook.AppointmentItem appointment = ResolveAppointment(control);
                if (appointment == null)
                {
                    // The stock add-in marks the element as clicked, invalidates it,
                    // and then awaits its asynchronous action. Yielding here gives
                    // Office time to repaint before the Inspector is created.
                    explorerMeetingActionActive = true;
                    LegacyTeamsSchedulerBridge.Log(
                        "Explorer meeting action queued; disabling controls");
                    InvalidateExplorerRibbon();
                    await Task.Delay(RibbonRepaintDelayMilliseconds);
                    appointment = (Outlook.AppointmentItem)outlook.CreateItem(
                        Outlook.OlItemType.olAppointmentItem);
                    appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                    appointment.Display(false);
                    ActivateAppointmentInspector(appointment);
                }
                if (String.IsNullOrWhiteSpace(invitationLanguage))
                    invitationLanguage = CleanRoomMeetingService.GetInvitationLanguage(
                        appointment);
                LegacyTeamsSchedulerBridge.Result meeting =
                    LegacyTeamsSchedulerBridge.CreateMeeting(appointment, templateId);
                CleanRoomMeetingService.ApplyMeeting(appointment,
                    meeting.MeetingId, meeting.JoinUrl,
                    meeting.BodyHtml, meeting.BodyText, meeting.OptionsUrl,
                    invitationLanguage);
                // The meeting properties drive CreateMeeting_GetVisible and
                // MeetingActions_GetVisible. Office does not reevaluate those
                // callbacks merely because the item properties changed, so refresh
                // the Inspector ribbon as soon as ApplyMeeting has completed.
                InvalidateRibbons();
                appointment.Display(false);
            }
            catch (Exception ex)
            {
                explorerMeetingActionActive = false;
                InvalidateExplorerRibbon();
                if (ex is OperationCanceledException) return;
                ShowError("Meeting creation failed", ex);
            }
        }

        private static void ActivateAppointmentInspector(
            Outlook.AppointmentItem appointment)
        {
            Outlook.Inspector inspector = null;
            try
            {
                // Display(false) can show an Inspector without making it active. In
                // that state Outlook keeps painting the Explorer ribbon until the
                // user clicks the editor. Explicit activation selects the correct
                // appointment ribbon as soon as the window becomes visible.
                inspector = appointment.GetInspector;
                if (inspector != null) inspector.Activate();
            }
            finally
            {
                ReleaseComObject(inspector);
            }
        }

        public void MeetNow(IRibbonControl control)
        {
            Outlook.AppointmentItem appointment = null;
            try
            {
                appointment = (Outlook.AppointmentItem)outlook.CreateItem(
                    Outlook.OlItemType.olAppointmentItem);
                appointment.Subject = "Réunion instantanée";
                appointment.Start = DateTime.Now;
                appointment.End = DateTime.Now.AddHours(1);
                LegacyTeamsSchedulerBridge.Result meeting =
                    LegacyTeamsSchedulerBridge.CreateMeetNow(appointment);
                Process.Start(meeting.JoinUrl);
                appointment.Close(Outlook.OlInspectorClose.olDiscard);
            }
            catch (Exception ex)
            {
                ShowError("Meet Now failed", ex);
            }
            finally
            {
                ReleaseComObject(appointment);
            }
        }

        private static string FormatException(Exception exception)
        {
            Exception current = exception;
            while (current.InnerException != null) current = current.InnerException;
            return current.GetType().FullName + "\r\n\r\n" + current.Message;
        }

        private static void ShowError(string operation, Exception exception)
        {
            LegacyTeamsSchedulerBridge.LogException(operation, exception);
            MessageBox.Show(FormatException(exception), DialogTitle,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void ReleaseComObject(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.ReleaseComObject(value); }
            catch (InvalidComObjectException) { }
        }

        public void ConnectOffice(IRibbonControl control)
        {
            try
            {
                OfficeNativeSignIn.Show();
                OfficeNativeSignIn.SetStandaloneEnabled(true);
                InvalidateRibbons();
            }
            catch (Exception ex)
            {
                if (ex is OperationCanceledException) return;
                MessageBox.Show(ex.Message, "Connexion Office native",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool ConnectOffice_GetVisible(IRibbonControl control)
        {
            return !IsStandaloneConnected();
        }

        public bool CreateMeeting_GetVisible(IRibbonControl control)
        {
            return IsStandaloneConnected() &&
                !CleanRoomMeetingService.HasMeeting(ResolveAppointment(control));
        }

        public bool ExplorerMeeting_GetVisible(IRibbonControl control)
        { return IsStandaloneConnected(); }

        public bool Language_GetVisible(IRibbonControl control)
        { return IsStandaloneConnected(); }

        public bool MeetingActions_GetVisible(IRibbonControl control)
        { return CleanRoomMeetingService.HasMeeting(ResolveAppointment(control)); }

        public void JoinMeeting(IRibbonControl control)
        {
            string url = CleanRoomMeetingService.GetJoinUrl(ResolveAppointment(control));
            if (!String.IsNullOrWhiteSpace(url)) Process.Start(url);
        }

        public void MeetingOptions(IRibbonControl control)
        {
            string url = CleanRoomMeetingService.GetOptionsUrl(ResolveAppointment(control));
            if (String.IsNullOrWhiteSpace(url))
                throw new InvalidOperationException("Teams n'a retourne aucune URL d'options.");
            Process.Start(url);
        }

        public void RemoveOnlineMeeting(IRibbonControl control)
        {
            Outlook.AppointmentItem appointment = ResolveAppointment(control);
            if (appointment != null) CleanRoomMeetingService.RemoveMeeting(appointment);
            InvalidateRibbons();
        }

        public bool OfficeAccount_GetVisible(IRibbonControl control)
        {
            return IsStandaloneConnected();
        }

        public string OfficeAccount_GetLabel(IRibbonControl control)
        {
            return OfficeNativeSignIn.GetAccountLabel();
        }

        public bool Ribbon_AlwaysDisabled(IRibbonControl control)
        {
            return false;
        }

        public void DisconnectStandalone(IRibbonControl control)
        {
            OfficeNativeSignIn.SetStandaloneEnabled(false);
            LegacyTeamsSchedulerBridge.Log(
                "Standalone account disconnected; Office identity unchanged");
            InvalidateRibbons();
        }

        private static bool IsStandaloneConnected()
        {
            return OfficeNativeSignIn.IsStandaloneEnabled() &&
                OfficeNativeSignIn.IsConnected();
        }

        private static Bitmap DrawFrenchFlag()
        {
            var bitmap = new Bitmap(32, 32);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillRectangle(Brushes.RoyalBlue, 2, 7, 9, 18);
                graphics.FillRectangle(Brushes.White, 11, 7, 9, 18);
                graphics.FillRectangle(Brushes.Red, 20, 7, 10, 18);
                graphics.DrawRectangle(Pens.Gray, 2, 7, 28, 18);
            }
            return bitmap;
        }

        private static Bitmap DrawUsFlag()
        {
            var bitmap = new Bitmap(32, 32);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Transparent);
                const float left = 1f;
                const float top = 6f;
                const float width = 30f;
                const float height = 20f;
                float stripeHeight = height / 13f;
                for (int row = 0; row < 13; row++)
                    graphics.FillRectangle(row % 2 == 0
                        ? Brushes.Firebrick : Brushes.White, left,
                        top + row * stripeHeight, width, stripeHeight + 0.2f);
                graphics.FillRectangle(Brushes.MidnightBlue, left, top,
                    13f, stripeHeight * 7f);
                for (int row = 0; row < 5; row++)
                    for (int column = 0; column < 6; column++)
                    {
                        float x = left + 1.2f + column * 2f + (row % 2 == 0 ? 0f : 1f);
                        float y = top + 1f + row * 1.8f;
                        graphics.FillEllipse(Brushes.White, x, y, 0.9f, 0.9f);
                    }
                graphics.DrawRectangle(Pens.Gray, left, top, width - 1f,
                    height - 1f);
            }
            return bitmap;
        }

        private Bitmap DrawAccountIcon()
        {
            var bitmap = new Bitmap(32, 32);
            Image profileImage = TryGetOutlookProfilePicture();
            string profilePath = null;
            if (profileImage == null)
                profilePath = OfficeNativeSignIn.GetProfilePicturePath();
            if (profileImage != null ||
                (!String.IsNullOrWhiteSpace(profilePath) && File.Exists(profilePath)))
            {
                try
                {
                    using (Image source = profileImage ?? new Bitmap(profilePath))
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    using (var clip = new System.Drawing.Drawing2D.GraphicsPath())
                    {
                        graphics.SmoothingMode =
                            System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        graphics.InterpolationMode =
                            System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphics.Clear(Color.Transparent);
                        clip.AddEllipse(2, 2, 28, 28);
                        graphics.SetClip(clip);
                        graphics.DrawImage(source, 2, 2, 28, 28);
                        graphics.ResetClip();
                        graphics.DrawEllipse(Pens.LightGray, 2, 2, 28, 28);
                        return bitmap;
                    }
                }
                catch (ArgumentException)
                {
                    // A partially refreshed Office cache image is non-fatal; the
                    // generic account glyph below remains available.
                }
                catch (IOException) { }
                catch (System.Runtime.InteropServices.ExternalException exception)
                {
                    LegacyTeamsSchedulerBridge.Log(
                        "Account image rendering failed: " + exception.ErrorCode);
                }
            }
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var background = new SolidBrush(Color.FromArgb(91, 95, 199)))
            using (var foreground = new SolidBrush(Color.White))
            {
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                graphics.FillEllipse(background, 2, 2, 28, 28);
                graphics.FillEllipse(foreground, 11, 7, 10, 10);
                graphics.FillEllipse(foreground, 7, 18, 18, 11);
            }
            return bitmap;
        }

        private Image TryGetOutlookProfilePicture()
        {
            Outlook.NameSpace session = null;
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry addressEntry = null;
            object exchangeUser = null;
            object nativePicture = null;
            try
            {
                if (outlook == null) return null;
                session = outlook.Session;
                currentUser = session.CurrentUser;
                addressEntry = currentUser == null ? null : currentUser.AddressEntry;
                if (addressEntry == null) return null;

                // ExchangeUser.GetPicture is the same identity-aware Outlook Object
                // Model path used for persona photos. Dynamic dispatch keeps the
                // build compatible with the pinned Outlook 15 PIA.
                dynamic dynamicAddressEntry = addressEntry;
                exchangeUser = dynamicAddressEntry.GetExchangeUser();
                if (exchangeUser == null) return null;
                dynamic dynamicExchangeUser = exchangeUser;
                nativePicture = dynamicExchangeUser.GetPicture();
                if (nativePicture == null) return null;
                Image converted = PictureConverter.FromComPicture(nativePicture);
                LegacyTeamsSchedulerBridge.Log(
                    "Account image loaded from Outlook ExchangeUser");
                if (converted == null) return null;
                using (converted) return new Bitmap(converted);
            }
            catch (COMException exception)
            {
                LegacyTeamsSchedulerBridge.Log(
                    "Outlook profile image unavailable: " + exception.ErrorCode);
                return null;
            }
            catch (Exception exception)
            {
                LegacyTeamsSchedulerBridge.LogException(
                    "Outlook profile image lookup failed", exception);
                return null;
            }
            finally
            {
                ReleaseComObject(nativePicture);
                ReleaseComObject(exchangeUser);
                ReleaseComObject(addressEntry);
                ReleaseComObject(currentUser);
                ReleaseComObject(session);
            }
        }

        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base(String.Empty) { }

            internal static Image FromComPicture(object picture)
            {
                return GetPictureFromIPicture(picture);
            }
        }

        private Outlook.AppointmentItem ResolveAppointment(IRibbonControl control)
        {
            // Ribbon callbacks always provide their owning Inspector as Context.
            // Falling back to Application.ActiveInspector used to create an extra
            // RCW and could also target the wrong window during focus transitions.
            object context = control == null ? null : control.Context;
            Outlook.Inspector inspector = context as Outlook.Inspector;
            if (inspector != null) return inspector.CurrentItem as Outlook.AppointmentItem;
            return null;
        }

        private const string ExplorerRibbon =
            "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='ExplorerRibbon_Load'>" +
            "<ribbon><tabs><tab idMso='TabCalendar'><group id='TmaCleanRoom.Calendar' label='TMA autonome' insertAfterMso='GroupCalendarNew'>" +
            "<button id='TmaCleanRoom.Connect' label='Se connecter' size='large' getImage='Ribbon_GetImage' getVisible='ConnectOffice_GetVisible' onAction='ConnectOffice'/>" +
            "<menu id='TmaCleanRoom.AccountMenu' label='Compte' size='large' getImage='Ribbon_GetImage' getVisible='OfficeAccount_GetVisible'>" +
            "<button id='TmaCleanRoom.AccountIdentity' getLabel='OfficeAccount_GetLabel' getImage='Ribbon_GetImage' getEnabled='Ribbon_AlwaysDisabled'/>" +
            "<menuSeparator id='TmaCleanRoom.AccountSeparator'/>" +
            "<button id='TmaCleanRoom.SwitchAccount' label='Changer de compte…' onAction='ConnectOffice'/>" +
            "<button id='TmaCleanRoom.Disconnect' label='Déconnecter TMA autonome' onAction='DisconnectStandalone'/>" +
            "</menu>" +
            "<button id='TmaCleanRoom.MeetNow' label='Réunion instantanée' size='large' keytip='MN' getImage='Ribbon_GetImage' getVisible='ExplorerMeeting_GetVisible' getEnabled='ExplorerMeeting_GetEnabled' onAction='MeetNow'/>" +
            "<splitButton id='TmaCleanRoom.MeetingSplit' size='large' getVisible='ExplorerMeeting_GetVisible' getEnabled='ExplorerMeeting_GetEnabled'>" +
            "<button id='TmaCleanRoom.Create' label='Réunion Teams' keytip='TM' getImage='Ribbon_GetImage' onAction='CreateMeeting'/>" +
            "<menu id='TmaCleanRoom.MeetingMenu'>" +
            "<button id='TmaCleanRoom.Schedule' label='Planifier une réunion' onAction='CreateMeeting'/>" +
            "<button id='TmaCleanRoom.Webinar' label='Webinaire' onAction='CreateWebinar'/>" +
            "<button id='TmaCleanRoom.TownHall' label='Assemblée' onAction='CreateTownHall'/>" +
            "<button id='TmaCleanRoom.VirtualAppointment' label='Rendez-vous virtuel' onAction='CreateVirtualAppointment'/>" +
            "</menu></splitButton>" +
            "</group></tab></tabs></ribbon></customUI>";

        private const string AppointmentRibbon =
            "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='AppointmentRibbon_Load'>" +
            "<ribbon><tabs><tab idMso='TabAppointment'><group id='TmaCleanRoom.Appointment' label='TMA autonome' insertAfterMso='GroupShow'>" +
            "<button id='TmaCleanRoom.Connect' label='Se connecter' size='large' getImage='Ribbon_GetImage' getVisible='ConnectOffice_GetVisible' onAction='ConnectOffice'/>" +
            "<menu id='TmaCleanRoom.LanguageMenu' label='Langue' size='large' getImage='Ribbon_GetImage' getVisible='Language_GetVisible'>" +
            "<button id='TmaCleanRoom.LanguageFr' label='Français' getImage='Ribbon_GetImage' onAction='SelectFrenchLanguage'/>" +
            "<button id='TmaCleanRoom.LanguageEn' label='English (US)' getImage='Ribbon_GetImage' onAction='SelectEnglishLanguage'/>" +
            "</menu>" +
            "<splitButton id='TmaCleanRoom.AppointmentMeetingSplit' size='large' getVisible='CreateMeeting_GetVisible'>" +
            "<button id='TmaCleanRoom.AppointmentCreate' label='Réunion Teams' keytip='TM' getImage='Ribbon_GetImage' onAction='CreateMeeting'/>" +
            "<menu id='TmaCleanRoom.AppointmentMeetingMenu'>" +
            "<button id='TmaCleanRoom.AppointmentSchedule' label='Planifier une réunion' onAction='CreateMeeting'/>" +
            "<button id='TmaCleanRoom.AppointmentWebinar' label='Webinaire' onAction='CreateWebinar'/>" +
            "<button id='TmaCleanRoom.AppointmentTownHall' label='Assemblée' onAction='CreateTownHall'/>" +
            "<button id='TmaCleanRoom.AppointmentVirtual' label='Rendez-vous virtuel' onAction='CreateVirtualAppointment'/>" +
            "</menu></splitButton>" +
            "<button id='TmaCleanRoom.Join' label='Rejoindre la réunion Teams' size='large' getImage='Ribbon_GetImage' getVisible='MeetingActions_GetVisible' onAction='JoinMeeting'/>" +
            "<button id='TmaCleanRoom.Options' label='Options de réunion' size='large' getImage='Ribbon_GetImage' getVisible='MeetingActions_GetVisible' onAction='MeetingOptions'/>" +
            "<button id='TmaCleanRoom.Remove' label='Ne pas héberger en ligne' size='large' imageMso='DeclineInvitation' getVisible='MeetingActions_GetVisible' onAction='RemoveOnlineMeeting'/>" +
            "<button id='TmaCleanRoom.Settings' label='Configuration' size='large' getImage='Ribbon_GetImage' getVisible='MeetingActions_GetVisible' onAction='MeetingOptions'/>" +
            "</group></tab></tabs></ribbon></customUI>";
    }
}
