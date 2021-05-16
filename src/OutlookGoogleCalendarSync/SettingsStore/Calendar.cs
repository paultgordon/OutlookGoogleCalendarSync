using log4net;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace OutlookGoogleCalendarSync.SettingsStore {
    [DataContract(Namespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync")]
    public class Calendar {
        private static readonly ILog log = LogManager.GetLogger(typeof(Calendar));

        public Calendar() {
            setDefaults();
        }

        //Default values before loading from xml and attribute not yet serialized
        [OnDeserializing]
        void OnDeserializing(StreamingContext context) {
            setDefaults();
        }

        private void setDefaults() {
            _ProfileName = "Default";

            //Outlook
            OutlookService = OutlookOgcs.Calendar.Service.DefaultMailbox;
            MailboxName = "";
            SharedCalendar = "";
            UseOutlookCalendar = new OutlookCalendarListEntry();
            CategoriesRestrictBy = RestrictBy.Exclude;
            Categories = new List<String>();
            OnlyRespondedInvites = false;
            OutlookDateFormat = "g";
            outlookGalBlocked = false;

            //Google
            UseGoogleCalendar = new GoogleCalendarListEntry();
            CloakEmail = true;
            ExcludeDeclinedInvites = true;
            ExcludeGoals = true;

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
            MaxAttendees = 200;
            AddColours = false;
            MergeItems = true;
            DisableDelete = true;
            ConfirmOnDelete = true;
            TargetCalendar = Sync.Direction.OutlookToGoogle;
            CreatedItemsOnly = true;
            SetEntriesPrivate = false;
            SetEntriesAvailable = false;
            AvailabilityStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olFree.ToString();
            SetEntriesColour = false;
            SetEntriesColourValue = Microsoft.Office.Interop.Outlook.OlCategoryColor.olCategoryColorNone.ToString();
            SetEntriesColourName = "None";
            SetEntriesColourGoogleId = "0";
            ColourMaps = new ColourMappingDictionary();
            Obfuscation = new Obfuscate();
            
            ExtirpateOgcsMetadata = false;
        }

        [DataMember] public string _ProfileName { get; set; }

        #region Outlook
        public enum RestrictBy {
            Include, Exclude
        }
        [DataMember] public OutlookOgcs.Calendar.Service OutlookService { get; set; }
        [DataMember] public string MailboxName { get; set; }
        [DataMember] public string SharedCalendar { get; set; }
        [DataMember] public OutlookCalendarListEntry UseOutlookCalendar { get; set; }
        [DataMember] public RestrictBy CategoriesRestrictBy { get; set; }
        [DataMember] public List<string> Categories { get; set; }
        /// <summary>Only allow Outlook to have one category assigned</summary>
        [DataMember] public Boolean SingleCategoryOnly { get; set; }
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
        [DataMember] public Boolean ExcludeDeclinedInvites { get; set; }
        [DataMember] public Boolean ExcludeGoals { get; set; }
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
        [DataMember] public int MaxAttendees { get; set; }
        [DataMember] public bool AddColours { get; set; }
        [DataMember] public bool MergeItems { get; set; }
        [DataMember] public bool DisableDelete { get; set; }
        [DataMember] public bool ConfirmOnDelete { get; set; }
        [DataMember] public Sync.Direction TargetCalendar { get; set; }
        [DataMember] public Boolean CreatedItemsOnly { get; set; }
        [DataMember] public bool SetEntriesPrivate { get; set; }
        [DataMember] public bool SetEntriesAvailable { get; set; }
        [DataMember] public String AvailabilityStatus { get; set; }
        [DataMember] public bool SetEntriesColour { get; set; }

        /// <summary>Set all Outlook appointments to this OlCategoryColor</summary>
        [DataMember] public String SetEntriesColourValue { get; set; }
        /// <summary>Set all Outlook appointments to this custom category name</summary>
        [DataMember] public String SetEntriesColourName { get; set; }
        /// <summary>Set all Google events to this colour ID</summary>
        [DataMember] public String SetEntriesColourGoogleId { get; set; }
        [DataMember]
        public ColourMappingDictionary ColourMaps { get; private set; }
        [CollectionDataContract(
            ItemName = "ColourMap",
            KeyName = "OutlookCategoryName",
            ValueName = "GoogleColourId",
            Namespace = "http://schemas.datacontract.org/2004/07/OutlookGoogleCalendarSync"
        )]
        public class ColourMappingDictionary : Dictionary<String, String> { }

        //Obfuscation
        [DataMember] public Obfuscate Obfuscation { get; set; }
        #endregion

        #region Advanced - Non GUI
        [DataMember] public Boolean ExtirpateOgcsMetadata { get; private set; }
        #endregion

        public void SetActive() {
            if (Settings.Instance.ActiveCalendarProfile != null &&
                Settings.Instance.ActiveCalendarProfile == this) return;

            log.Debug("Changing active settings profile '" + this._ProfileName + "'.");
            Settings.Instance.ActiveCalendarProfile = this;

            if (Forms.Main.Instance.Visible) 
                Forms.Main.Instance?.UpdateGUIsettings_Profile();
        }
    }
}
