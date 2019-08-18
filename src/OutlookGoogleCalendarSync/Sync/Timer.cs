using log4net;
using System;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync.Sync {
    public class SyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(SyncTimer));
        private Timer ogcsTimer;
        
        /// <summary>
        /// Should only be set via SettingsStore class
        /// </summary>
        public DateTime LastSyncDate { internal get; set; }

        private DateTime nextSyncDate;
        public DateTime NextSyncDate {
            get { return nextSyncDate; }
            set {
                nextSyncDate = value;
                NextSyncDateText = nextSyncDate.ToLongDateString() + " @ " + nextSyncDate.ToLongTimeString();
                if (Forms.Main.Instance.ActiveCalendarProfile.OutlookPush) NextSyncDateText += " + Push";
                Forms.Main.Instance.NextSyncVal = NextSyncDateText;
                log.Info("Next sync scheduled for " + NextSyncDateText);
            }
        }
        public String NextSyncDateText { get; internal set; }
        
        public SyncTimer(DateTime lastSync) {
            ogcsTimer = new Timer();
            this.Tag = "AutoSyncTimer";
            this.Tick += new EventHandler(ogcsTimer_Tick);

            //Refresh synchronizations (last and next)
            this.LastSyncDate = lastSync;
            SetNextSync(getResyncInterval());
        }

        private void ogcsTimer_Tick(object sender, EventArgs e) {
            if (Forms.ErrorReporting.Instance.Visible) return;
            log.Debug("Scheduled sync triggered.");

            Forms.Main frm = Forms.Main.Instance;
            frm.NotificationTray.ShowBubbleInfo("Autosyncing calendars: " + Forms.Main.Instance.ActiveCalendarProfile.SyncDirection.Name + "...");
            if (!Sync.Engine.Instance.SyncingNow) {
                frm.Sync_Click(sender, null);
            } else {
                log.Debug("Busy syncing already. Rescheduled for 5 mins time.");
                SetNextSync(5, fromNow: true);
            }
        }

        private int getResyncInterval() {
            int min = Forms.Main.Instance.ActiveCalendarProfile.SyncInterval;
            if (Forms.Main.Instance.ActiveCalendarProfile.SyncIntervalUnit == "Hours") {
                min *= 60;
            }
            return min;
        }

        public void SetNextSync(int? delayMins = null, Boolean fromNow = false) {
            int _delayMins = delayMins ?? getResyncInterval();

            if (Forms.Main.Instance.ActiveCalendarProfile.SyncInterval != 0) {
                DateTime nextSyncDate = this.LastSyncDate.AddMinutes(_delayMins);
                DateTime now = DateTime.Now;
                if (fromNow)
                    nextSyncDate = now.AddMinutes(_delayMins);

                if (this.Interval != (delayMins * 60000)) {
                    this.Stop();
                    TimeSpan diff = nextSyncDate - now;
                    if (diff.TotalMinutes < 1) {
                        nextSyncDate = now.AddMinutes(1);
                        this.Interval = 1 * 60000;
                    } else {
                        this.Interval = (int)(diff.TotalMinutes * 60000);
                    }
                    this.Start();
                }
                NextSyncDate = nextSyncDate;
            } else {
                this.Stop();
                Forms.Main.Instance.NextSyncVal = this.Status();
                log.Info("Schedule disabled.");
            }
        }

        public void Switch(Boolean enable) {
            if (enable && !this.Enabled) this.Start();
            else if (!enable && this.Enabled) this.Stop();
        }

        public Boolean Running() {
            return this.Enabled;
        }

        public String Status() {
            if (this.Running()) return NextSyncDateText;
            else if (Forms.Main.Instance.ActiveCalendarProfile.OgcsPushTimer != null && Forms.Main.Instance.ActiveCalendarProfile.OgcsPushTimer.Running()) return "Push Sync Active";
            else return "Inactive";
        }
    }


    public class PushSyncTimer : Timer {
        private static readonly ILog log = LogManager.GetLogger(typeof(PushSyncTimer));
        private Timer ogcsTimer;
        private DateTime lastRunTime;
        private Int32 lastRunItemCount;
        private Int16 failures = 0;
        private static PushSyncTimer instance;
        public static PushSyncTimer Instance {
            get {
                if (instance == null) {
                    instance = new PushSyncTimer();
                }
                return instance;
            }
        }

        private PushSyncTimer() {
            ResetLastRun();
            ogcsTimer = new Timer();
            this.Tag = "PushTimer";
            this.Interval = 2 * 60000;
            this.Tick += new EventHandler(ogcsPushTimer_Tick);
        }

        /// <summary>
        /// Recalculate item count as of now.
        /// </summary>
        public void ResetLastRun() {
            this.lastRunTime = DateTime.Now;
            try {
                log.Fine("Updating calendar item count following Push Sync.");
                this.lastRunItemCount = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(true).Count;
            } catch {
                log.Error("Failed to update item count following a Push Sync.");
            }
        }

        private void ogcsPushTimer_Tick(object sender, EventArgs e) {
            if (Forms.ErrorReporting.Instance.Visible) return;
            log.Fine("Push sync triggered.");

            try {
                System.Collections.Generic.List<Microsoft.Office.Interop.Outlook.AppointmentItem> items = OutlookOgcs.Calendar.Instance.GetCalendarEntriesInRange(true);

                if (items.Count < this.lastRunItemCount || items.FindAll(x => x.LastModificationTime > this.lastRunTime).Count > 0) {
                    log.Debug("Changes found for Push sync.");
                    Forms.Main.Instance.NotificationTray.ShowBubbleInfo("Autosyncing calendars: " + Forms.Main.Instance.ActiveCalendarProfile.SyncDirection.Name + "...");
                    if (!Sync.Engine.Instance.SyncingNow) {
                        Forms.Main.Instance.Sync_Click(sender, null);
                    } else {
                        log.Debug("Busy syncing already. No need to push.");
                    }
                } else {
                    log.Fine("No changes found.");
                }
                failures = 0;
            } catch (System.Exception ex) {
                failures++;
                OGCSexception.Analyse("Push Sync failed " + failures + " times to check for changed items.", ex);
                if (failures == 10)
                    Forms.Main.Instance.Console.UpdateWithError("Push Sync is failing.", ex, notifyBubble: true);
            }
        }

        public void Switch(Boolean enable) {
            if (enable && !this.Enabled) {
                ResetLastRun();
                this.Start();
                if (Forms.Main.Instance.ActiveCalendarProfile.SyncInterval == 0) Forms.Main.Instance.NextSyncVal = "Push Sync Active";
            } else if (!enable && this.Enabled) {
                this.Stop();
                Forms.Main.Instance.ActiveCalendarProfile.OgcsTimer.SetNextSync();
            }
        }
        public Boolean Running() {
            return this.Enabled;
        }
    }
}
