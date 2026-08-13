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
        private Outlook.Application outlook;
        private IRibbonUI ribbon;
        private Outlook.Inspectors inspectors;
        private readonly List<Outlook.Inspector> appointmentInspectors =
            new List<Outlook.Inspector>();
        private Timer inspectorCloseTimer;

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            outlook = application as Outlook.Application;
            if (outlook != null)
            {
                inspectors = outlook.Inspectors;
                inspectors.NewInspector += Inspectors_NewInspector;
                for (int index = 1; index <= inspectors.Count; index++)
                    TrackAppointmentInspector(inspectors[index]);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            if (inspectors != null)
                inspectors.NewInspector -= Inspectors_NewInspector;
            foreach (Outlook.Inspector inspector in appointmentInspectors)
            {
                try { ((Outlook.InspectorEvents_10_Event)inspector).Close -= AppointmentInspector_Close; }
                catch { }
                Marshal.FinalReleaseComObject(inspector);
            }
            appointmentInspectors.Clear();
            if (inspectors != null) Marshal.FinalReleaseComObject(inspectors);
            inspectors = null;
            if (outlook != null) Marshal.FinalReleaseComObject(outlook);
            outlook = null;
            ribbon = null;
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

        public void Ribbon_Load(IRibbonUI ui) { ribbon = ui; }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            TrackAppointmentInspector(inspector);
            if (ribbon != null) ribbon.Invalidate();
        }

        private void TrackAppointmentInspector(Outlook.Inspector inspector)
        {
            if (inspector == null) return;
            object item = inspector.CurrentItem;
            if (!(item is Outlook.AppointmentItem)) return;
            ((Outlook.InspectorEvents_10_Event)inspector).Close += AppointmentInspector_Close;
            appointmentInspectors.Add(inspector);
        }

        private void AppointmentInspector_Close()
        {
            if (inspectorCloseTimer != null) inspectorCloseTimer.Dispose();
            inspectorCloseTimer = new Timer { Interval = 300 };
            inspectorCloseTimer.Tick += delegate
            {
                inspectorCloseTimer.Stop();
                inspectorCloseTimer.Dispose();
                inspectorCloseTimer = null;
                CleanupClosedAppointmentInspectors();
                if (ribbon != null) ribbon.Invalidate();
            };
            inspectorCloseTimer.Start();
        }

        private void CleanupClosedAppointmentInspectors()
        {
            for (int index = appointmentInspectors.Count - 1; index >= 0; index--)
            {
                Outlook.Inspector inspector = appointmentInspectors[index];
                bool closed;
                try { closed = inspector.CurrentItem == null; }
                catch (COMException) { closed = true; }
                if (!closed) continue;
                try { ((Outlook.InspectorEvents_10_Event)inspector).Close -= AppointmentInspector_Close; }
                catch { }
                appointmentInspectors.RemoveAt(index);
                Marshal.FinalReleaseComObject(inspector);
            }
        }

        public bool ExplorerMeeting_GetEnabled(IRibbonControl control)
        {
            if (inspectors == null) return true;
            for (int index = 1; index <= inspectors.Count; index++)
            {
                Outlook.Inspector inspector;
                try
                {
                    inspector = inspectors[index];
                    object item = inspector.CurrentItem;
                    if (item is Outlook.AppointmentItem) return false;
                }
                catch (COMException) { }
            }
            return true;
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
            if (ribbon != null) ribbon.Invalidate();
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

        private void CreateMeetingWithTemplate(IRibbonControl control, string templateId,
            string invitationLanguage)
        {
            try
            {
                Outlook.AppointmentItem appointment = ResolveAppointment(control);
                if (appointment == null)
                {
                    appointment = (Outlook.AppointmentItem)outlook.CreateItem(
                        Outlook.OlItemType.olAppointmentItem);
                    appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
                    appointment.Display(false);
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
                appointment.Display(false);
            }
            catch (Exception ex)
            {
                if (ex is OperationCanceledException) return;
                MessageBox.Show(FormatException(ex), "TMA Clean Room",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show(FormatException(ex), "TMA Clean Room",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (appointment != null)
                    Marshal.FinalReleaseComObject(appointment);
            }
        }

        private static string FormatException(Exception exception)
        {
            Exception current = exception;
            while (current.InnerException != null) current = current.InnerException;
            return current.GetType().FullName + "\r\n\r\n" + current.Message;
        }

        public void ConnectOffice(IRibbonControl control)
        {
            try
            {
                OfficeNativeSignIn.Show();
                if (ribbon != null) ribbon.Invalidate();
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
            return !OfficeNativeSignIn.IsConnected();
        }

        public bool CreateMeeting_GetVisible(IRibbonControl control)
        {
            return OfficeNativeSignIn.IsConnected() &&
                !CleanRoomMeetingService.HasMeeting(ResolveAppointment(control));
        }

        public bool ExplorerMeeting_GetVisible(IRibbonControl control)
        { return OfficeNativeSignIn.IsConnected(); }

        public bool Language_GetVisible(IRibbonControl control)
        { return OfficeNativeSignIn.IsConnected(); }

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
            if (ribbon != null) ribbon.Invalidate();
        }

        public bool OfficeAccount_GetVisible(IRibbonControl control)
        {
            return OfficeNativeSignIn.IsConnected();
        }

        public string OfficeAccount_GetLabel(IRibbonControl control)
        {
            return "Connecte : " + OfficeNativeSignIn.GetAccountLabel();
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

        private Outlook.AppointmentItem ResolveAppointment(IRibbonControl control)
        {
            object context = control == null ? null : control.Context;
            Outlook.Inspector inspector = context as Outlook.Inspector;
            if (inspector != null) return inspector.CurrentItem as Outlook.AppointmentItem;
            if (outlook == null) return null;
            Outlook.Inspector active = outlook.ActiveInspector();
            return active == null ? null : active.CurrentItem as Outlook.AppointmentItem;
        }

        private const string ExplorerRibbon =
            "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>" +
            "<ribbon><tabs><tab idMso='TabCalendar'><group id='TmaCleanRoom.Calendar' label='TMA autonome' insertAfterMso='GroupCalendarNew'>" +
            "<button id='TmaCleanRoom.Connect' label='Se connecter' size='large' getImage='Ribbon_GetImage' getVisible='ConnectOffice_GetVisible' onAction='ConnectOffice'/>" +
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
            "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>" +
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
