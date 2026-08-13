using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Globalization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TmaCleanRoom
{
    internal static class CleanRoomMeetingService
    {
        private const string PropertyBase =
            "http://schemas.microsoft.com/mapi/string/{9A5B5D75-42A7-4EF1-98BE-8B4B70944E89}/";
        internal const string MeetingIdProperty = PropertyBase + "TmaCleanRoomMeetingId";
        internal const string JoinUrlProperty = PropertyBase + "TmaCleanRoomJoinUrl";
        internal const string SchemaProperty = PropertyBase + "TmaCleanRoomSchema";
        internal const string OptionsUrlProperty = PropertyBase + "TmaCleanRoomOptionsUrl";
        internal const string InvitationLanguageProperty =
            PropertyBase + "TmaCleanRoomInvitationLanguage";

        internal static void MarkPrototype(Outlook.AppointmentItem appointment)
        {
            appointment.PropertyAccessor.SetProperty(SchemaProperty, "1");
            appointment.Save();
        }

        internal static void EnableOutlookOnlineMeeting(
            Outlook.AppointmentItem appointment)
        {
            if (appointment == null) throw new ArgumentNullException("appointment");

            // These members exist on the current Outlook COM object model but are
            // absent from the Outlook 15 PIA installed in the GAC. Late binding is
            // intentional: Outlook/Exchange remains responsible for authentication,
            // provider selection and creation of the actual online meeting.
            dynamic nativeAppointment = appointment;
            nativeAppointment.GetOnlineMeetingProvider();
            int provider = Convert.ToInt32(nativeAppointment.OnlineMeetingProvider);
            if (provider == 0)
            {
                throw new InvalidOperationException(
                    "Outlook ne trouve aucun fournisseur de reunion en ligne pour ce compte. " +
                    "Verifiez que le compte Microsoft 365 est connecte dans Outlook.");
            }

            appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            nativeAppointment.IsOnlineMeeting = true;
            appointment.Save();
            appointment.Display(false);
        }

        internal static void ApplyMeeting(Outlook.AppointmentItem appointment,
            string meetingId, string joinUrl, string stockHtml, string stockText,
            string optionsUrl, string invitationLanguage)
        {
            if (String.IsNullOrWhiteSpace(meetingId) || String.IsNullOrWhiteSpace(joinUrl))
                throw new ArgumentException("Meeting ID and join URL are required.");
            appointment.PropertyAccessor.SetProperty(MeetingIdProperty, meetingId);
            appointment.PropertyAccessor.SetProperty(JoinUrlProperty, joinUrl);
            appointment.PropertyAccessor.SetProperty(SchemaProperty, "1");
            if (!String.IsNullOrWhiteSpace(optionsUrl))
                appointment.PropertyAccessor.SetProperty(OptionsUrlProperty, optionsUrl);
            string existingBody = appointment.Body ?? String.Empty;
            string plainBlock = !String.IsNullOrWhiteSpace(stockText) ? stockText :
                "____________________________________________________________\r\n" +
                "Reunion Microsoft Teams\r\n\r\n" +
                "Rejoindre la reunion maintenant :\r\n" + joinUrl + "\r\n" +
                "ID de reunion : " + meetingId + "\r\n" +
                "____________________________________________________________\r\n\r\n";

            try
            {
                string meetingHtml = BuildMeetingHtml(meetingId, joinUrl,
                    optionsUrl, stockText, stockHtml, invitationLanguage);
                LegacyTeamsSchedulerBridge.Log("Invitation template built: chars=" +
                    meetingHtml.Length + ", customMarker=" +
                    meetingHtml.Contains("data-tma-clean-room=\"meeting\""));
                InsertHtmlWithWordEditor(appointment, meetingHtml);
                LegacyTeamsSchedulerBridge.Log("Invitation template inserted with WordEditor");
            }
            catch (Exception exception)
            {
                LegacyTeamsSchedulerBridge.LogException(
                    "Invitation template insertion failed; using plain text", exception);
                // Older Outlook object models may not expose HTMLBody for an
                // appointment. Keep a readable plain-text fallback.
                appointment.Body = plainBlock + existingBody;
            }
        }

        internal static string GetJoinUrl(Outlook.AppointmentItem appointment)
        { return GetStringProperty(appointment, JoinUrlProperty); }

        internal static string GetOptionsUrl(Outlook.AppointmentItem appointment)
        { return GetStringProperty(appointment, OptionsUrlProperty); }

        internal static bool HasMeeting(Outlook.AppointmentItem appointment)
        { return !String.IsNullOrWhiteSpace(GetJoinUrl(appointment)); }

        internal static string GetInvitationLanguage(Outlook.AppointmentItem appointment)
        { return GetStringProperty(appointment, InvitationLanguageProperty); }

        internal static void SetInvitationLanguage(Outlook.AppointmentItem appointment,
            string language)
        {
            if (appointment == null) return;
            appointment.PropertyAccessor.SetProperty(InvitationLanguageProperty,
                language);
            if (!HasMeeting(appointment)) return;

            string meetingId = GetStringProperty(appointment, MeetingIdProperty);
            string joinUrl = GetJoinUrl(appointment);
            string optionsUrl = GetOptionsUrl(appointment);
            string currentText = appointment.Body;
            RemoveMeetingBody(appointment, joinUrl);
            ApplyMeeting(appointment, meetingId, joinUrl, null, currentText,
                optionsUrl, language);
        }

        internal static void RemoveMeeting(Outlook.AppointmentItem appointment)
        {
            string joinUrl = GetJoinUrl(appointment);
            RemoveMeetingBody(appointment, joinUrl);
            DeleteProperty(appointment, MeetingIdProperty);
            DeleteProperty(appointment, JoinUrlProperty);
            DeleteProperty(appointment, OptionsUrlProperty);
            DeleteProperty(appointment, SchemaProperty);
            DeleteProperty(appointment, InvitationLanguageProperty);
        }

        private static void RemoveMeetingBody(Outlook.AppointmentItem appointment,
            string joinUrl)
        {
            try
            {
                Outlook.Inspector inspector = appointment.GetInspector;
                dynamic document = inspector.WordEditor;
                dynamic links = document.Hyperlinks;
                int start = Int32.MaxValue;
                int end = -1;
                for (int i = 1; i <= links.Count; i++)
                {
                    dynamic link = links.Item(i);
                    if (String.Equals(Convert.ToString(link.Address), joinUrl,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        int linkStart = Convert.ToInt32(link.Range.Start);
                        start = Math.Min(start, linkStart);
                        dynamic paragraph = link.Range.Paragraphs.Item(1);
                        end = Math.Max(end, Convert.ToInt32(paragraph.Range.End));
                    }
                }
                if (end >= 0)
                {
                    dynamic tables = document.Tables;
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        dynamic table = tables.Item(i);
                        int tableStart = Convert.ToInt32(table.Range.Start);
                        int tableEnd = Convert.ToInt32(table.Range.End);
                        if (tableStart <= start && tableEnd >= start)
                        {
                            start = Math.Min(start, tableStart);
                            end = Math.Max(end, tableEnd);
                        }
                    }
                    dynamic range = document.Range(start, end);
                    range.Delete();
                }
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inspector);
            }
            catch { }
        }

        private static string GetStringProperty(Outlook.AppointmentItem appointment,
            string property)
        {
            if (appointment == null) return null;
            try { return Convert.ToString(appointment.PropertyAccessor.GetProperty(property)); }
            catch { return null; }
        }

        private static void DeleteProperty(Outlook.AppointmentItem appointment,
            string property)
        {
            try { appointment.PropertyAccessor.DeleteProperty(property); }
            catch { }
        }

        private static void InsertHtmlWithWordEditor(
            Outlook.AppointmentItem appointment, string html)
        {
            string temporaryFile = Path.Combine(Path.GetTempPath(),
                "TmaCleanRoom-" + Guid.NewGuid().ToString("N") + ".html");
            try
            {
                string document = "<!DOCTYPE html><html><head>" +
                    "<meta charset=\"utf-8\"></head><body>" + html +
                    "</body></html>";
                File.WriteAllText(temporaryFile, document,
                    new System.Text.UTF8Encoding(false));
                Outlook.Inspector inspector = appointment.GetInspector;
                if (inspector == null)
                    throw new InvalidOperationException(
                        "Outlook ne fournit aucun inspecteur pour inserer l'invitation.");
                try
                {
                    dynamic wordDocument = inspector.WordEditor;
                    dynamic range = wordDocument.Content;
                    string existingText = Convert.ToString(range.Text);
                    range.Collapse(1); // Word.WdCollapseDirection.wdCollapseStart
                    if (!String.IsNullOrWhiteSpace(existingText == null ? null :
                        existingText.Trim('\r', '\n', '\a', ' ')))
                    {
                        // Outlook a déjà inséré la signature au moment de Display().
                        // Le bloc de réunion est ajouté avant celle-ci sans la remplacer.
                        range.InsertParagraphBefore();
                        range.Collapse(1);
                    }
                    // Word defaults Attachment to true when optional arguments are
                    // omitted. With a dynamic COM dispatch, values must be passed
                    // without C# ref modifiers; ref makes the runtime binder reject
                    // FileName before Word even receives the call.
                    range.InsertFile(temporaryFile, Type.Missing,
                        false, false, false);
                    if (String.IsNullOrWhiteSpace(existingText == null ? null :
                        existingText.Trim('\r', '\n', '\a', ' ')))
                    {
                        InsertConfiguredOutlookSignature(wordDocument);
                    }
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(inspector);
                }
            }
            finally
            {
                if (File.Exists(temporaryFile)) File.Delete(temporaryFile);
            }
        }

        private static void InsertConfiguredOutlookSignature(dynamic wordDocument)
        {
            try
            {
                string signatureName = Convert.ToString(wordDocument.Application
                    .EmailOptions.EmailSignature.NewMessageSignature);
                if (String.IsNullOrWhiteSpace(signatureName)) return;
                string signaturePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Signatures", signatureName + ".htm");
                if (!File.Exists(signaturePath)) return;
                dynamic signatureRange = wordDocument.Content;
                signatureRange.Collapse(0);
                signatureRange.InsertParagraphAfter();
                signatureRange.Collapse(0);
                signatureRange.InsertFile(signaturePath, Type.Missing,
                    false, false, false);
                LegacyTeamsSchedulerBridge.Log("Outlook signature inserted: " + signatureName);
            }
            catch (Exception exception)
            {
                LegacyTeamsSchedulerBridge.LogException(
                    "Outlook signature insertion failed", exception);
            }
        }

        private static string BuildMeetingHtml(string meetingId, string joinUrl,
            string optionsUrl, string stockText, string stockHtml,
            string invitationLanguage)
        {
            string language = String.IsNullOrWhiteSpace(invitationLanguage)
                ? CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
                : invitationLanguage;
            bool french = String.Equals(language, "fr",
                StringComparison.OrdinalIgnoreCase);
            string safeId = HtmlEncode(meetingId);
            string safeUrl = HtmlEncode(joinUrl);
            string iconDataUri = LoadTeamsIconDataUri();
            string passcode = ExtractPasscode(stockText, stockHtml);
            string passcodeRow = String.IsNullOrWhiteSpace(passcode) ? String.Empty :
                "<tr><td style=\"padding:5px 0;color:#616161;width:120px\">" +
                (french ? "Code secret" : "Passcode") + "</td>" +
                "<td style=\"padding:5px 0;color:#242424;font-weight:600\">" +
                "<span class=\"MsoNoProof\" style=\"mso-no-proof:yes;" +
                "text-decoration:none;color:#242424\">" +
                HtmlEncode(passcode) + "</span></td></tr>";
            string optionsBlock = String.IsNullOrWhiteSpace(optionsUrl) ? String.Empty :
                "<a href=\"" + HtmlEncode(optionsUrl) +
                "\" style=\"color:#5b5fc7;text-decoration:underline;font-weight:600\">" +
                (french ? "Options de réunion" : "Meeting options") + "</a>" +
                "<span style=\"color:#b3b3b3;padding:0 8px\">|</span>";
            string assemblyDirectory = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);
            string templateName = french ? "MeetingInvite.html" :
                "MeetingInvite.en-US.html";
            string templatePath = Path.Combine(assemblyDirectory, "Templates",
                templateName);
            LegacyTeamsSchedulerBridge.Log("Invitation template path: " + templatePath);
            if (!File.Exists(templatePath))
                throw new FileNotFoundException(
                    "Le template HTML de l'invitation Teams est introuvable.", templatePath);
            string template = File.ReadAllText(templatePath);
            if (template.IndexOf("{{JOIN_URL}}", StringComparison.Ordinal) < 0 ||
                template.IndexOf("{{MEETING_ID}}", StringComparison.Ordinal) < 0 ||
                template.IndexOf("{{TEAMS_ICON}}", StringComparison.Ordinal) < 0)
                throw new InvalidDataException(
                    "Le template HTML ne contient pas les marqueurs requis.");
            return template.Replace("{{JOIN_URL}}", safeUrl)
                .Replace("{{MEETING_ID}}", safeId)
                .Replace("{{TEAMS_ICON}}", iconDataUri)
                .Replace("{{PASSCODE_ROW}}", passcodeRow)
                .Replace("{{OPTIONS_BLOCK}}", optionsBlock);
        }

        private static string LoadTeamsIconDataUri()
        {
            string assemblyDirectory = Path.GetDirectoryName(
                Assembly.GetExecutingAssembly().Location);
            string iconPath = Path.Combine(assemblyDirectory, "Assets",
                "NewMeeting_Large_96.png");
            if (!File.Exists(iconPath)) return String.Empty;
            return "data:image/png;base64," +
                Convert.ToBase64String(File.ReadAllBytes(iconPath));
        }

        private static string ExtractPasscode(string stockText, string stockHtml)
        {
            string source = !String.IsNullOrWhiteSpace(stockText) ? stockText : stockHtml;
            if (String.IsNullOrWhiteSpace(source)) return null;
            Match match = Regex.Match(source,
                @"(?:Passcode|Code\s+secret|Mot\s+de\s+passe)\s*:?\s*([^\s<]+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? match.Groups[1].Value.Trim() : null;
        }

        private static string HtmlEncode(string value)
        {
            if (String.IsNullOrEmpty(value)) return String.Empty;
            return value.Replace("&", "&amp;").Replace("<", "&lt;")
                .Replace(">", "&gt;").Replace("\"", "&quot;")
                .Replace("'", "&#39;");
        }
    }
}
