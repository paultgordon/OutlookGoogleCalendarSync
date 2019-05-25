﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Xml.Linq;

namespace OutlookGoogleCalendarSync.SettingsStore {
    public static class Upgrade {
        private static readonly ILog log = LogManager.GetLogger(typeof(Upgrade));

        //OGCS releases that require the settings XML to be upgraded
        private const Int32 multipleCalendars = 2070901; //v2.7.9.1;

        private static String settingsVersion;
        private static Int32 settingsVersionNum;


        public static void Check() {
            settingsVersion = XMLManager.ImportElement("Version", Settings.ConfigFile);
            settingsVersionNum = Program.VersionToInt(settingsVersion);

            while (upgradePerformed()) {
            }
        }


        private static Boolean upgradePerformed() {
            try {
                if (settingsVersionNum < multipleCalendars) {
                    upgradeToMultiCalendar();
                    settingsVersionNum = multipleCalendars;
                    return true;
                } else
                    return false;
            } catch {
                log.Warn("Upgrade(s) didn't complete successfully. The user will likely need to reset their settings.");
                return false;
            }
        }

        private static void backupSettingsFile() {
            try {
                log.Info("Backing up '" + Settings.ConfigFile + "' for v" + settingsVersion);
                String backupFile = System.Text.RegularExpressions.Regex.Replace(Settings.ConfigFile, @"(\.\w+)$", "-v" + settingsVersion + "$1");
                File.Copy(Settings.ConfigFile, backupFile);
                log.Info(backupFile + " created.");
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Failed to create backup settings file", ex);
            }
        }

        private static void upgradeToMultiCalendar() {
            backupSettingsFile();

            XDocument xml = null;
            try {
                xml = XDocument.Load(Settings.ConfigFile);
                XElement settingsElement = XMLManager.GetElement("Settings", xml);
                XElement calendarsElement = XMLManager.AddElement("Calendars", settingsElement);
                XElement calendarElement = XMLManager.AddElement("Calendar", calendarsElement);

                XMLManager.MoveElement("OutlookService", settingsElement, calendarElement);
                XMLManager.MoveElement("MailboxName", settingsElement, calendarElement);
                XMLManager.MoveElement("SharedCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("UseOutlookCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("CategoriesRestrictBy", settingsElement, calendarElement);
                XMLManager.MoveElement("Categories", settingsElement, calendarElement);
                XMLManager.MoveElement("OnlyRespondedInvites", settingsElement, calendarElement);
                XMLManager.MoveElement("OutlookDateFormat", settingsElement, calendarElement);
                XMLManager.MoveElement("outlookGalBlocked", settingsElement, calendarElement);
                
                XMLManager.MoveElement("UseGoogleCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("CloakEmail", settingsElement, calendarElement);

                XMLManager.MoveElement("SyncDirection", settingsElement, calendarElement);
                XMLManager.MoveElement("DaysInThePast", settingsElement, calendarElement);
                XMLManager.MoveElement("DaysInTheFuture", settingsElement, calendarElement);
                XMLManager.MoveElement("SyncInterval", settingsElement, calendarElement);
                XMLManager.MoveElement("SyncIntervalUnit", settingsElement, calendarElement);
                XMLManager.MoveElement("OutlookPush", settingsElement, calendarElement);
                XMLManager.MoveElement("AddLocation", settingsElement, calendarElement);
                XMLManager.MoveElement("AddDescription", settingsElement, calendarElement);
                XMLManager.MoveElement("AddDescription_OnlyToGoogle", settingsElement, calendarElement);
                XMLManager.MoveElement("AddReminders", settingsElement, calendarElement);
                XMLManager.MoveElement("UseGoogleDefaultReminder", settingsElement, calendarElement);
                XMLManager.MoveElement("UseOutlookDefaultReminder", settingsElement, calendarElement);
                XMLManager.MoveElement("ReminderDND", settingsElement, calendarElement);
                XMLManager.MoveElement("ReminderDNDstart", settingsElement, calendarElement);
                XMLManager.MoveElement("ReminderDNDend", settingsElement, calendarElement);
                XMLManager.MoveElement("AddAttendees", settingsElement, calendarElement);
                XMLManager.MoveElement("AddColours", settingsElement, calendarElement);
                XMLManager.MoveElement("MergeItems", settingsElement, calendarElement);
                XMLManager.MoveElement("DisableDelete", settingsElement, calendarElement);
                XMLManager.MoveElement("ConfirmOnDelete", settingsElement, calendarElement);
                XMLManager.MoveElement("TargetCalendar", settingsElement, calendarElement);
                XMLManager.MoveElement("CreatedItemsOnly", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesPrivate", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesAvailable", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesColour", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesColourValue", settingsElement, calendarElement);
                XMLManager.MoveElement("SetEntriesColourName", settingsElement, calendarElement);
                XMLManager.MoveElement("Obfuscation", settingsElement, calendarElement);
                
                XMLManager.MoveElement("ExtirpateOgcsMetadata", settingsElement, calendarElement);
                 
            } catch (System.Exception ex) {
                OGCSexception.Analyse("Problem encountered whilst upgrading " + Settings.ConfigFilename, ex);
                throw ex;
            } finally {
                if (xml != null) {
                    xml.Root.Sort();
                    try {
                        xml.Save(Settings.ConfigFile);
                    } catch (System.Exception ex) {
                        OGCSexception.Analyse("Could not save upgraded settings file " + Settings.ConfigFile, ex);
                        throw ex;
                    }
                }
            }
        }
    }
}