﻿using System;
using System.Runtime.Serialization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookGoogleCalendarSync.SettingsStore {
    [DataContract(Namespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync")]
    public class Calendar {

        public Calendar() {
            setDefaults();
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        private void setDefaults() {
            //Outlook
            OutlookService = OutlookOgcs.Calendar.Service.DefaultMailbox;
            MailboxName = "";
            SharedCalendar = "";
            UseOutlookCalendar = new OutlookCalendarListEntry();
            CategoriesRestrictBy = RestrictBy.Exclude;
            Categories = new System.Collections.Generic.List<String>();
            OnlyRespondedInvites = false;
            OutlookDateFormat = "g";
            outlookGalBlocked = false;

            //Google
            UseGoogleCalendar = new GoogleCalendarListEntry();
            CloakEmail = true;

            //Sync Options
            SyncDirection = Sync.Direction.OutlookToGoogle;
            DaysInThePast = 1;
            DaysInTheFuture = 60;
            SyncInterval = 0;
            SyncIntervalUnit = "Hours";
            OutlookPush = false;
            AddLocation = true;
            AddDescription = true;
            AddDescription_OnlyToGoogle = true;
            AddReminders = false;
            UseGoogleDefaultReminder = false;
            UseOutlookDefaultReminder = false;
            ReminderDND = false;
            ReminderDNDstart = DateTime.Now.Date.AddHours(22);
            ReminderDNDend = DateTime.Now.Date.AddDays(1).AddHours(6);
            AddAttendees = false;
            AddColours = false;
            MergeItems = true;
            DisableDelete = true;
            ConfirmOnDelete = true;
            TargetCalendar = Sync.Direction.OutlookToGoogle;
            CreatedItemsOnly = true;
            SetEntriesPrivate = false;
            SetEntriesAvailable = false;
            SetEntriesColour = false;
            SetEntriesColourValue = Microsoft.Office.Interop.Outlook.OlCategoryColor.olCategoryColorNone.ToString();
            SetEntriesColourName = "None";
            Obfuscation = new Obfuscate();
            
            ExtirpateOgcsMetadata = false;
        }

        #region Outlook
        public enum RestrictBy {
            Include, Exclude
        }
        [DataMember] public OutlookOgcs.Calendar.Service OutlookService { get; set; }
        [DataMember] public string MailboxName { get; set; }
        [DataMember] public string SharedCalendar { get; set; }
        [DataMember] public OutlookCalendarListEntry UseOutlookCalendar { get; set; }
        [DataMember] public RestrictBy CategoriesRestrictBy { get; set; }
        [DataMember] public System.Collections.Generic.List<string> Categories { get; set; }
        [DataMember] public Boolean OnlyRespondedInvites { get; set; }
        [DataMember] public string OutlookDateFormat { get; set; }
        private Boolean outlookGalBlocked;
        [DataMember] public Boolean OutlookGalBlocked {
            get { return outlookGalBlocked; }
            set {
                outlookGalBlocked = value;
                if (!Settings.Instance.Loading() && Forms.Main.Instance.IsHandleCreated) Forms.Main.Instance.FeaturesBlockedByCorpPolicy(value);
            }
        }
        #endregion
        #region Google
        [DataMember] public GoogleCalendarListEntry UseGoogleCalendar { get; set; }
        [DataMember] public Boolean CloakEmail { get; set; }
        #endregion
        #region Sync Options
        //Main
        public DateTime SyncStart { get { return DateTime.Today.AddDays(-DaysInThePast); } }
        public DateTime SyncEnd { get { return DateTime.Today.AddDays(+DaysInTheFuture + 1); } }
        [DataMember] public Sync.Direction SyncDirection { get; set; }
        [DataMember] public int DaysInThePast { get; set; }
        [DataMember] public int DaysInTheFuture { get; set; }
        [DataMember] public int SyncInterval { get; set; }
        [DataMember] public String SyncIntervalUnit { get; set; }
        [DataMember] public bool OutlookPush { get; set; }
        [DataMember] public bool AddLocation { get; set; }
        [DataMember] public bool AddDescription { get; set; }
        [DataMember] public bool AddDescription_OnlyToGoogle { get; set; }
        [DataMember] public bool AddReminders { get; set; }
        [DataMember] public bool UseGoogleDefaultReminder { get; set; }
        [DataMember] public bool UseOutlookDefaultReminder { get; set; }
        [DataMember] public bool ReminderDND { get; set; }
        [DataMember] public DateTime ReminderDNDstart { get; set; }
        [DataMember] public DateTime ReminderDNDend { get; set; }
        [DataMember] public bool AddAttendees { get; set; }
        [DataMember] public bool AddColours { get; set; }
        [DataMember] public bool MergeItems { get; set; }
        [DataMember] public bool DisableDelete { get; set; }
        [DataMember] public bool ConfirmOnDelete { get; set; }
        [DataMember] public Sync.Direction TargetCalendar { get; set; }
        [DataMember] public Boolean CreatedItemsOnly { get; set; }
        [DataMember] public bool SetEntriesPrivate { get; set; }
        [DataMember] public bool SetEntriesAvailable { get; set; }
        [DataMember] public bool SetEntriesColour { get; set; }
        [DataMember] public String SetEntriesColourValue { get; set; }
        [DataMember] public String SetEntriesColourName { get; set; }

        //Obfuscation
        [DataMember] public Obfuscate Obfuscation { get; set; }
        #endregion

        #region Advanced - Non GUI
        [DataMember] public Boolean ExtirpateOgcsMetadata { get; private set; }
        #endregion
    }
}