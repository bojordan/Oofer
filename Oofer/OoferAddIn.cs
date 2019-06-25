using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;

namespace Oofer
{
    /// <summary>
    /// https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-perform-actions-when-an-e-mail-message-is-received?view=vs-2019
    /// https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-send-e-mail-programmatically?view=vs-2019
    /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.meetingitem?view=outlook-pia
    /// </summary>
    public partial class OoferAddIn
    {
        private NameSpace _outlookNameSpace = null;
        private MAPIFolder _inbox = null;
        private Items _items = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _outlookNameSpace = this.Application.GetNamespace("MAPI");
            _inbox = _outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            _items = _inbox.Items;
            _items.ItemAdd += new ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
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

                    if ((apptItem.Subject.ToUpperInvariant().Contains("OOF") || apptItem.Subject.ToUpperInvariant().Contains("WFH")) &&
                        organizerAddressEntry != currentUserAddressEntry)
                    {
                        if (apptItem.ReminderSet || apptItem.BusyStatus != OlBusyStatus.olFree)
                        {
                            apptItem.BusyStatus = OlBusyStatus.olFree;
                            apptItem.ReminderSet = false;
                            apptItem.ResponseRequested = false;
                            apptItem.Save();

                            CreateEmailItem(
                                subjectEmail: $"Cleaned OOF/WFH: {apptItem.Subject}",
                                toEmail: currentUserAddressEntry.Address,
                                bodyEmail: $"This appointment was cleaned:\n\n{apptItem.Subject}");
                        }
                    }
                }
            }
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
