using Microsoft.Office.Interop.Outlook;
using System;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;

namespace Oofer
{
    /// <summary>
    /// https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-perform-actions-when-an-e-mail-message-is-received?view=vs-2019
    /// https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-send-e-mail-programmatically?view=vs-2019
    /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.meetingitem?view=outlook-pia
    /// https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-retrieve-unread-messages-from-the-inbox?view=vs-2019
    /// https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-automatically-accept-a-meeting-request
    /// </summary>
    public partial class OoferAddIn
    {
        private NameSpace _outlookNameSpace = null;
        private MAPIFolder _inbox = null;
        private Items _items = null;

        private string[] _matches = new[]
        {
            "OOF",
            "WFH",
            "DOCTOR",
            "OFFLINE",
            "DR ",
            " APPT",
            "APPOINTMENT",
            "TRAINING",
            "SICK",
            "VACATION",
            "WATCHING",
            "VIEWING",
            "ATTENDING"
        };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _outlookNameSpace = this.Application.GetNamespace("MAPI");
            _inbox = _outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            _items = _inbox.Items;
            _items.ItemAdd += new ItemsEvents_ItemAddEventHandler(Items_ItemAdd);

            // run on unread items at startup
            var unreadItems = _items.Restrict("[Unread]=true");

            foreach (var unreadItem in unreadItems)
            {
                Items_ItemAdd(unreadItem);
            }
        }

        private void Items_ItemAdd(object item)
        {
            MeetingItem meetingItem = item as MeetingItem;
            if (meetingItem != null)
            {
                var apptItem = meetingItem.GetAssociatedAppointment(false);
                if (apptItem != null)
                {
                    var organizerAddressEntry = apptItem.GetOrganizer();
                    var currentUserAddressEntry = apptItem.Session?.CurrentUser?.AddressEntry;

                    if (_matches.Any(x => apptItem.Subject.ToUpperInvariant().Contains(x)) &&
                        organizerAddressEntry != currentUserAddressEntry)
                    {
                        if (apptItem.ReminderSet || apptItem.BusyStatus != OlBusyStatus.olFree)
                        {
                            var subjectText = apptItem.Subject;
                            var busyStatusText = GetBusyStatusString(apptItem.BusyStatus);
                            var reminderSetText = apptItem.ReminderSet ? "Reminder was set" : "Reminder was not set";
                            var responseRequestedText = apptItem.ResponseRequested ? "Response was requested" : "Response was not requested";

                            apptItem.BusyStatus = OlBusyStatus.olFree;
                            apptItem.ReminderSet = false;
                            apptItem.ResponseRequested = false;
                            
                            // We will not send the response, but want to accept the appointment on our side. See linked
                            // docs to send the response, if desired.
                            apptItem.Respond(OlMeetingResponse.olMeetingAccepted, true, Type.Missing);
                            apptItem.UnRead = false;

                            apptItem.Save();

                            meetingItem.UnRead = false;
                            meetingItem.Save();

                            CreateEmailItem(
                                subjectEmail: $"Cleaned OOF/WFH: {apptItem.Subject}",
                                toEmail: currentUserAddressEntry.Address,
                                bodyEmail: $"This appointment was cleaned:\n\n{subjectText}\n{busyStatusText}\n{reminderSetText}\n{responseRequestedText}");
                        }
                    }
                }
            }
        }

        private static string GetBusyStatusString(OlBusyStatus busyStatus)
        {
            switch (busyStatus)
            {
                case OlBusyStatus.olBusy:
                    return "Busy";
                case OlBusyStatus.olFree:
                    return "Free";
                case OlBusyStatus.olOutOfOffice:
                    return "Out of office";
                case OlBusyStatus.olTentative:
                    return "Tentative";
                case OlBusyStatus.olWorkingElsewhere:
                    return "Working elsewhere";
            }

            return "Unknown";
        }

        private void CreateEmailItem(string subjectEmail, string toEmail, string bodyEmail)
        {
            MailItem eMail = (MailItem) this.Application.CreateItem(OlItemType.olMailItem);
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Importance = OlImportance.olImportanceLow;
            eMail.Send();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
