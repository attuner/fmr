/**
 * The Big Family Radio Backend - Google Apps Script
 * Studio Master High-Fidelity Audio Stream & Broadcast Sync Engine
 */

const FOLDER_NAME = "Feel Music Radio";
const ADMIN_SECRET_KEY = "admin123";
const DEFAULT_STATION_NAME = "FM Radio";

function doGet(e) {
  try {
    e = e || { parameter: { action: "getStationData" } };
    const parameter = e.parameter || {};
    const action = parameter.action || "getStationData";

    if (action === "streamAudio") {
      const fileId = parameter.fileId;
      if (!fileId) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Missing fileId" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      try {
        const file = DriveApp.getFileById(fileId);
        let mime = file.getMimeType() || "audio/mpeg";
        if (mime === "application/octet-stream") mime = "audio/mpeg";
        
        const bytes = file.getBlob().getBytes();
        const b64 = Utilities.base64Encode(bytes);
        const dataUri = "data:" + mime + ";base64," + b64;
        
        return ContentService.createTextOutput(JSON.stringify({ 
          success: true, 
          dataUri: dataUri,
          mimeType: mime,
          sizeBytes: bytes.length
        })).setMimeType(ContentService.MimeType.JSON);
      } catch(err) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    let response = { success: false, message: "Invalid request" };

    if (action === "getStationData") {
      response = getStationData();
    } else if (action === "getAdminData") {
      if (parameter.key === ADMIN_SECRET_KEY) {
        response = getAdminData();
      } else {
        response = { success: false, message: "Unauthorized admin access" };
      }
    }

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "No post data received" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let contents = {};
    try {
      contents = JSON.parse(e.postData.contents);
    } catch(parseErr) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Invalid JSON body" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const action = contents.action;
    let response = { success: false, message: "Invalid action" };

    if (action === "register") response = handleRegister(contents);
    else if (action === "login") response = handleLogin(contents);
    else if (action === "uploadAudio") response = handleAudioUpload(contents);
    else if (action === "adminLogin") response = handleAdminLogin(contents);
    else if (action === "adminUpdateUserStatus") response = handleAdminUpdateUser(contents);
    else if (action === "adminUpdateAudioApproval") response = handleUpdateAudioApproval(contents);
    else if (action === "adminDeleteTrack") response = handleDeleteTrack(contents);
    else if (action === "adminUpdateTrackCategory") response = handleUpdateTrackCategory(contents);
    else if (action === "adminSaveSettings") response = handleSaveSettings(contents);
    else if (action === "heartbeat") response = handleHeartbeat(contents);
    else if (action === "addScheduledItem") response = handleAddScheduledItem(contents);
    else if (action === "deleteScheduledItem") response = handleDeleteScheduledItem(contents);
    else if (action === "adminPushTrack") response = handlePushTrack(contents);
    else if (action === "adminPushLiveMic") response = handlePushLiveMic(contents);

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------- SHEET & DRIVE HELPERS -----------------

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(FOLDER_NAME);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("No active spreadsheet bound to Apps Script.");

  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["Username", "Mobile", "Password", "Status", "RegisteredAt"]);
  }

  let tracksSheet = ss.getSheetByName("Tracks");
  const trackHeaders = ["Seq", "Title", "Description", "Contributor", "FileId", "StreamUrl", "CreatedAt", "Category", "ApprovalStatus", "DurationSec"];
  if (!tracksSheet) {
    tracksSheet = ss.insertSheet("Tracks");
    tracksSheet.appendRow(trackHeaders);
  } else {
    const lastCol = tracksSheet.getLastColumn();
    if (lastCol < trackHeaders.length) {
      tracksSheet.getRange(1, 1, 1, trackHeaders.length).setValues([trackHeaders]);
    }
  }

  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.appendRow(["Key", "Value"]);
    settingsSheet.appendRow(["station_name", DEFAULT_STATION_NAME]);
    settingsSheet.appendRow(["station_tagline", "Live Synchronized Family Broadcast"]);
    settingsSheet.appendRow(["seq_version", String(Date.now())]);
    settingsSheet.appendRow(["pushed_track", ""]);
    settingsSheet.appendRow(["push_version", ""]);
  }

  let listenersSheet = ss.getSheetByName("Listeners");
  if (!listenersSheet) {
    listenersSheet = ss.insertSheet("Listeners");
    listenersSheet.appendRow(["ListenerId", "Username", "Status", "LastSeen"]);
  }

  let scheduledSheet = ss.getSheetByName("ScheduledTimeline");
  if (!scheduledSheet) {
    scheduledSheet = ss.insertSheet("ScheduledTimeline");
    scheduledSheet.appendRow(["ScheduleId", "ExactTime", "SlotName", "FileId", "CreatedAt"]);
  }

  return { usersSheet, tracksSheet, settingsSheet, listenersSheet, scheduledSheet };
}

// ----------------- USER AUTH & UPLOAD -----------------

function handleRegister(data) {
  data = data || {};
  const username = String(data.username || "").trim();
  const mobile = String(data.mobile || "").trim();
  const password = String(data.password || "").trim();

  if (!username || !mobile || !password) return { success: false, message: "All fields are required" };

  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === mobile) {
      return { success: false, message: "Mobile number already registered" };
    }
  }

  usersSheet.appendRow([username, mobile, password, "pending", new Date().toISOString()]);
  return { success: true, message: "Registration successful. Welcome to the Family!", user: { username, mobile, status: "pending" } };
}

function handleLogin(data) {
  data = data || {};
  const mobile = String(data.mobile || "").trim();
  const password = String(data.password || "").trim();

  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === mobile && String(rows[i][2]).trim() === password) {
      return { success: true, user: { username: String(rows[i][0] || "Family Member"), mobile: String(rows[i][1] || ""), status: String(rows[i][3] || "pending") } };
    }
  }
  return { success: false, message: "Invalid credentials" };
}

function handleAudioUpload(data) {
  data = data || {};
  const mobile = String(data.mobile || "").trim();
  const username = String(data.username || "Family Member").trim();
  const title = String(data.title || "Family Audio").trim();
  const description = String(data.description || "").trim();
  const base64File = data.base64File;
  const fileName = data.fileName || `${title}.mp3`;
  const mimeType = data.mimeType || "audio/mpeg";
  const durationSec = Math.round(Number(data.durationSec)) || 140;
  const isAdmin = data.adminKey === ADMIN_SECRET_KEY;
  const initialCategory = String(data.category || (isAdmin ? "Random Plays" : "Random Plays")).trim();

  if (!base64File) return { success: false, message: "Missing audio data" };

  const { usersSheet, tracksSheet, settingsSheet } = getSheets();

  if (!isAdmin) {
    const userRows = usersSheet.getDataRange().getValues();
    let isApproved = false;
    for (let i = 1; i < userRows.length; i++) {
      if (String(userRows[i][1]).trim() === mobile && String(userRows[i][3]).trim() === "approved") {
        isApproved = true;
        break;
      }
    }
    if (!isApproved) return { success: false, message: "Admin approval is required to upload." };
  }

  const folder = getOrCreateFolder();
  const decodedData = Utilities.base64Decode(base64File);
  const blob = Utilities.newBlob(decodedData, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const streamUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
  const createdAt = new Date().toISOString();
  const approvalStatus = isAdmin ? "approved" : "pending";

  const trackRows = tracksSheet.getDataRange().getValues();
  const nextSeq = trackRows.length;

  tracksSheet.appendRow([nextSeq, title, description, username, fileId, streamUrl, createdAt, initialCategory, approvalStatus, durationSec]);
  setSettingValue(settingsSheet, "seq_version", String(Date.now()));

  return {
    success: true,
    message: "Audio uploaded in original master quality!",
    track: { seq: nextSeq, title, description, fileId, streamUrl, createdAt, category: initialCategory, approvalStatus, durationSec }
  };
}

// ----------------- SCHEDULED PLAY TIMINGS HANDLERS -----------------

function handleAddScheduledItem(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const exactTime = String(data.exactTime || "").trim(); // "HH:mm" 24h format
  const fileId = String(data.fileId || "").trim();
  const slotName = String(data.slotName || "Special Broadcast").trim();

  if (!exactTime || !fileId) return { success: false, message: "Exact time and audio must be selected" };

  const { scheduledSheet, settingsSheet } = getSheets();
  const schedId = "SCH_" + Date.now();
  scheduledSheet.appendRow([schedId, exactTime, slotName, fileId, new Date().toISOString()]);

  const newVersion = String(Date.now());
  setSettingValue(settingsSheet, "seq_version", newVersion);

  return { success: true, message: `Audio scheduled for ${exactTime}!`, version: newVersion };
}

function handleDeleteScheduledItem(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const scheduleId = String(data.scheduleId || "").trim();

  const { scheduledSheet, settingsSheet } = getSheets();
  const rows = scheduledSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === scheduleId) {
      scheduledSheet.deleteRow(i + 1);
      const newVersion = String(Date.now());
      setSettingValue(settingsSheet, "seq_version", newVersion);
      return { success: true, message: "Scheduled audio removed." };
    }
  }
  return { success: false, message: "Schedule item not found" };
}

function getScheduledTimelineList() {
  const { scheduledSheet } = getSheets();
  const rows = scheduledSheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      list.push({
        scheduleId: String(rows[i][0]),
        exactTime: String(rows[i][1]),
        slotName: String(rows[i][2] || "Scheduled Broadcast"),
        fileId: String(rows[i][3]),
        createdAt: String(rows[i][4] || "")
      });
    }
  }

  list.sort((a, b) => a.exactTime.localeCompare(b.exactTime));
  return list;
}

function handleHeartbeat(data) {
  data = data || {};
  const listenerId = String(data.listenerId || "L_anon").trim();
  const username = String(data.username || "Family Member").trim();
  const isPlaying = Boolean(data.isPlaying);

  const { listenersSheet } = getSheets();
  const rows = listenersSheet.getDataRange().getValues();
  const now = Date.now();
  let found = false;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === listenerId) {
      listenersSheet.getRange(i + 1, 2).setValue(username);
      listenersSheet.getRange(i + 1, 3).setValue(isPlaying ? "Listening Live" : "Idle");
      listenersSheet.getRange(i + 1, 4).setValue(now);
      found = true;
      break;
    }
  }

  if (!found) {
    listenersSheet.appendRow([listenerId, username, isPlaying ? "Listening Live" : "Idle", now]);
  }

  let activeCount = 0;
  const updatedRows = listenersSheet.getDataRange().getValues();
  for (let i = 1; i < updatedRows.length; i++) {
    const lastSeen = Number(updatedRows[i][3]) || 0;
    if (now - lastSeen < 50000 && String(updatedRows[i][2]) === "Listening Live") {
      activeCount++;
    }
  }

  return { success: true, activeCount };
}

function getActiveListenersList() {
  const { listenersSheet } = getSheets();
  const rows = listenersSheet.getDataRange().getValues();
  const now = Date.now();
  const activeListeners = [];

  for (let i = 1; i < rows.length; i++) {
    const lastSeen = Number(rows[i][3]) || 0;
    if (now - lastSeen < 50000) {
      activeListeners.push({
        listenerId: String(rows[i][0] || ""),
        username: String(rows[i][1] || "Family Guest"),
        status: String(rows[i][2] || "Idle"),
        lastSeen: new Date(lastSeen).toLocaleTimeString()
      });
    }
  }
  return activeListeners;
}

function handlePushLiveMic(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  if (!data.base64File) return { success: false, message: "Missing audio recording" };

  const { tracksSheet, settingsSheet } = getSheets();
  const folder = getOrCreateFolder();

  const decodedData = Utilities.base64Decode(data.base64File);
  const blob = Utilities.newBlob(decodedData, "audio/webm", `Family_Announcement_${Date.now()}.webm`);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const streamUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
  const nextSeq = tracksSheet.getDataRange().getValues().length;
  const recordedDuration = Math.round(Number(data.durationSec)) || 60;

  tracksSheet.appendRow([nextSeq, "🔴 Live Family Announcement", "Urgent Family Message", "Family Admin", fileId, streamUrl, new Date().toISOString(), "Scheduled Playlist", "approved", recordedDuration]);

  const pushVer = String(Date.now());
  setSettingValue(settingsSheet, "pushed_track", fileId);
  setSettingValue(settingsSheet, "push_version", pushVer);

  return { success: true, message: "Live announcement broadcasted!", fileId, pushVersion: pushVer };
}

function handlePushTrack(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { settingsSheet } = getSheets();
  const fileId = String(data.fileId || "");
  const pushVer = String(Date.now());

  setSettingValue(settingsSheet, "pushed_track", fileId);
  setSettingValue(settingsSheet, "push_version", pushVer);

  return { success: true, message: "Audio pushed live across the radio!", pushVersion: pushVer };
}

function handleUpdateAudioApproval(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { tracksSheet, settingsSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();
  const targetId = String(data.fileId || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === targetId) {
      tracksSheet.getRange(i + 1, 9).setValue(data.status);
      setSettingValue(settingsSheet, "seq_version", String(Date.now()));
      return { success: true, message: `Status updated to ${data.status}` };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleUpdateTrackCategory(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { tracksSheet, settingsSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();
  const targetId = String(data.fileId || "").trim();
  const newCat = (data.category === "Scheduled Playlist") ? "Scheduled Playlist" : "Random Plays";

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === targetId) {
      tracksSheet.getRange(i + 1, 8).setValue(newCat);
      setSettingValue(settingsSheet, "seq_version", String(Date.now()));
      return { success: true, message: "Category updated: " + newCat };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleDeleteTrack(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { tracksSheet, scheduledSheet, settingsSheet } = getSheets();
  const rows = tracksSheet.getDataRange().getValues();
  const targetId = String(data.fileId || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][4]).trim() === targetId) {
      try { DriveApp.getFileById(targetId).setTrashed(true); } catch (err) {}
      tracksSheet.deleteRow(i + 1);

      // Clean up from scheduled table if present
      const schedRows = scheduledSheet.getDataRange().getValues();
      for (let s = schedRows.length - 1; s >= 1; s--) {
        if (String(schedRows[s][3]).trim() === targetId) {
          scheduledSheet.deleteRow(s + 1);
        }
      }

      setSettingValue(settingsSheet, "seq_version", String(Date.now()));
      return { success: true, message: "Audio removed from library" };
    }
  }
  return { success: false, message: "Track not found" };
}

function handleSaveSettings(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { settingsSheet } = getSheets();
  const s = data.settings || {};
  const rows = settingsSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const key = rows[i][0];
    if (s[key] !== undefined) {
      settingsSheet.getRange(i + 1, 2).setValue(s[key]);
    }
  }
  return { success: true, message: "Settings saved" };
}

function getStationData() {
  const { tracksSheet, settingsSheet } = getSheets();
  const trackRows = tracksSheet.getDataRange().getValues();
  const tracks = [];

  for (let i = 1; i < trackRows.length; i++) {
    if (trackRows[i][4]) {
      tracks.push({
        seq: Number(trackRows[i][0]) || i,
        title: String(trackRows[i][1] || "Untitled Audio"),
        description: String(trackRows[i][2] || ""),
        contributor: String(trackRows[i][3] || "Family"),
        fileId: String(trackRows[i][4]),
        streamUrl: String(trackRows[i][5]),
        createdAt: String(trackRows[i][6] || ""),
        category: String(trackRows[i][7] || "Random Plays"),
        approvalStatus: String(trackRows[i][8] || "pending"),
        durationSec: Math.round(Number(trackRows[i][9])) || 140
      });
    }
  }

  tracks.sort((a, b) => new Date(a.createdAt || 0) - new Date(b.createdAt || 0));

  const settingsRows = settingsSheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < settingsRows.length; i++) {
    settings[settingsRows[i][0]] = settingsRows[i][1];
  }

  const scheduled = getScheduledTimelineList();
  const activeListeners = getActiveListenersList();
  const liveCount = activeListeners.filter(l => l.status === "Listening Live").length;

  return {
    success: true,
    serverTime: Date.now(),
    stationName: settings.station_name || DEFAULT_STATION_NAME,
    stationTagline: settings.station_tagline || "Live Synchronized Family Broadcast",
    tracks,
    settings,
    scheduled,
    liveListenersCount: liveCount
  };
}

function handleAdminLogin(data) {
  data = data || {};
  if (data.key === ADMIN_SECRET_KEY) return { success: true, message: "Authenticated" };
  return { success: false, message: "Invalid Passcode" };
}

function getAdminData() {
  const { usersSheet } = getSheets();
  const station = getStationData();

  const userRows = usersSheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < userRows.length; i++) {
    if (userRows[i][0] || userRows[i][1]) {
      users.push({
        username: String(userRows[i][0] || "User"),
        mobile: String(userRows[i][1] || ""),
        status: String(userRows[i][3] || "pending"),
        registeredAt: String(userRows[i][4] || "")
      });
    }
  }

  return {
    success: true,
    serverTime: Date.now(),
    users,
    tracks: station.tracks || [],
    settings: station.settings || {},
    scheduled: station.scheduled || [],
    activeListeners: station.activeListeners || getActiveListenersList()
  };
}

function handleAdminUpdateUser(data) {
  data = data || {};
  if (data.adminKey !== ADMIN_SECRET_KEY) return { success: false, message: "Unauthorized" };
  const { usersSheet } = getSheets();
  const rows = usersSheet.getDataRange().getValues();
  const targetMobile = String(data.mobile || "").trim();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === targetMobile) {
      usersSheet.getRange(i + 1, 4).setValue(data.status);
      return { success: true, message: `Member status updated: ${data.status}` };
    }
  }
  return { success: false, message: "User not found" };
}

function setSettingValue(sheet, key, value) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}