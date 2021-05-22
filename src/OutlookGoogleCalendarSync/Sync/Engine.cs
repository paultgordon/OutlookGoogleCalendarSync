using Google.Apis.Calendar.v3.Data;
using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using log4net;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleCalendarSync.Sync {
    public partial class Engine {
        private static readonly ILog log = LogManager.GetLogger(typeof(Engine));

        private static Engine instance;
        public static Engine Instance {
            get {
                if (instance == null) instance = new Engine();
                return instance;
            }
            set {
                instance = value;
            }
        }

        public Engine() { }

        /// <summary>
        /// The profile currently set to be synced, either manually from GUI settings or scheduled from a Timer.
        /// </summary>
        public Object ActiveProfile;

        /// <summary>
        /// Get the earliest upcoming sync time
        /// </summary>
        public DateTime? NextSyncDate { get {
                DateTime? retVal = null;
                foreach (SettingsStore.Calendar cal in Settings.Instance.Calendars) {
                    retVal = cal.OgcsTimer.NextSyncDate < (retVal ?? DateTime.MaxValue) ? cal.OgcsTimer.NextSyncDate : retVal;
                }
                return retVal;
            }
        }

        /// <summary>The time the current sync started</summary>
        public DateTime SyncStarted { get; set; }

        public AbortableBackgroundWorker bwSync { get; private set; }
        public Boolean SyncingNow {
            get {
                if (bwSync == null) return false;
                else return bwSync.IsBusy;
            }
        }
        public Boolean CancellationPending {
            get {
                return (bwSync != null && bwSync.CancellationPending);
            }
        }
        public Boolean ManualForceCompare = false;
        public enum SyncResult {
            OK,
            Fail,
            Abandon,
            AutoRetry,
            ReconnectThenRetry,
            UserCancelled
        }

        public void Sync_Requested(object sender = null, EventArgs e = null) {
            ManualForceCompare = false;
            if (sender != null && sender.GetType().ToString().EndsWith("Timer")) { //Automated sync
                Forms.Main.Instance.NotificationTray.UpdateItem("delayRemove", enabled: false);
                if (Forms.Main.Instance.bSyncNow.Text == "Start Sync") {
                    Timer aTimer = sender as Timer;
                    log.Info("Scheduled sync started (" + aTimer.Tag.ToString() + ").");
                    if (aTimer.Tag.ToString() == "PushTimer") Start(updateSyncSchedule: false);
                    else if (aTimer.Tag.ToString() == "AutoSyncTimer") Sync.Engine.Instance.Start(updateSyncSchedule: true);
                } else if (Forms.Main.Instance.bSyncNow.Text == "Stop Sync") {
                    log.Warn("Automated sync triggered whilst previous sync is still running. Ignoring this new request.");
                    if (this.bwSync == null)
                        //May be inbetween setting button to "Stop Sync" and actually starting background worker
                        log.Debug("Background worker is null, sync in process of initialising?");
                    else
                        log.Debug("Background worker is busy? A:" + bwSync.IsBusy.ToString());
                }

            } else { //Manual sync
                if (Forms.Main.Instance.bSyncNow.Text == "Start Sync" || Forms.Main.Instance.bSyncNow.Text == "Start Full Sync") {
                    log.Info("Manual sync requested.");
                    if (SyncingNow) {
                        log.Info("Already busy syncing, cannot accept another sync request.");
                        MessageBox.Show("A sync is already running. Please wait for it to complete and then try again.", "Sync already running", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        return;
                    }
                    Sync.Engine.Instance.ActiveProfile = Forms.Main.Instance.ActiveCalendarProfile;
                    if (Control.ModifierKeys == Keys.Shift) {
                        if (Forms.Main.Instance.ActiveCalendarProfile.SyncDirection == Direction.Bidirectional) {
                            MessageBox.Show("Forcing a full sync is not allowed whilst in 2-way sync mode.\r\nPlease temporarily chose a direction to sync in first.",
                                "2-way full sync not allowed", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                        log.Info("Shift-click has forced a compare of all items");
                        ManualForceCompare = true;
                    }
                    Start(updateSyncSchedule: false);

                } else if (Forms.Main.Instance.bSyncNow.Text == "Stop Sync") {
                    GoogleOgcs.Calendar.Instance.Authenticator.CancelTokenSource.Cancel();
                    if (!SyncingNow) return;

                    if (!bwSync.CancellationPending) {
                        Forms.Main.Instance.Console.Update("Sync cancellation requested.", Console.Markup.warning);
                        bwSync.CancelAsync();
                    } else {
                        Forms.Main.Instance.Console.Update("Repeated cancellation requested - forcefully aborting sync!", Console.Markup.warning);
                        try {
                            bwSync.Abort();
                            bwSync.Dispose();
                            bwSync = null;
                        } catch { }
                    }
                }
            }
        }

        public void Start(Boolean updateSyncSchedule = true) {
            if (Settings.GetProfileType(this.ActiveProfile) == Settings.ProfileType.Calendar) {
                Sync.Engine.Calendar calendar = new Sync.Engine.Calendar(this.ActiveProfile as SettingsStore.Calendar);
                calendar.StartSync(updateSyncSchedule);
            }
        }

        #region Compare Event Attributes
        public static Boolean CompareAttribute(String attrDesc, Direction fromTo, String googleAttr, String outlookAttr, StringBuilder sb, ref int itemModified) {
            if (googleAttr == null) googleAttr = "";
            if (outlookAttr == null) outlookAttr = "";
            //Truncate long strings
            String googleAttr_stub = ((googleAttr.Length > 50) ? googleAttr.Substring(0, 47) + "..." : googleAttr).Replace("\r\n", " ");
            String outlookAttr_stub = ((outlookAttr.Length > 50) ? outlookAttr.Substring(0, 47) + "..." : outlookAttr).Replace("\r\n", " ");
            log.Fine("Comparing " + attrDesc);
            log.UltraFine("Google  attribute: " + googleAttr);
            log.UltraFine("Outlook attribute: " + outlookAttr);
            if (googleAttr != outlookAttr) {
                if (fromTo == Direction.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr_stub + " => " + googleAttr_stub);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr_stub + " => " + outlookAttr_stub);
                }
                itemModified++;
                log.Fine("Attributes differ.");
                return true;
            }
            return false;
        }
        public static Boolean CompareAttribute(String attrDesc, Direction fromTo, Boolean googleAttr, Boolean outlookAttr, StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing " + attrDesc);
            log.UltraFine("Google  attribute: " + googleAttr);
            log.UltraFine("Outlook attribute: " + outlookAttr);
            if (googleAttr != outlookAttr) {
                if (fromTo == Direction.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr + " => " + googleAttr);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr + " => " + outlookAttr);
                }
                itemModified++;
                log.Fine("Attributes differ.");
                return true;
            }
            return false;
        }
        public static Boolean CompareAttribute(String attrDesc, Direction fromTo, DateTime googleAttr, DateTime outlookAttr, StringBuilder sb, ref int itemModified) {
            log.Fine("Comparing " + attrDesc);
            log.UltraFine("Google  attribute: " + googleAttr);
            log.UltraFine("Outlook attribute: " + outlookAttr);
            if (googleAttr != outlookAttr) {
                if (fromTo == Direction.GoogleToOutlook) {
                    sb.AppendLine(attrDesc + ": " + outlookAttr + " => " + googleAttr);
                } else {
                    sb.AppendLine(attrDesc + ": " + googleAttr + " => " + outlookAttr);
                }
                itemModified++;
                log.Fine("Attributes differ.");
                return true;
            }
            return false;
        }
        #endregion
    }
}
