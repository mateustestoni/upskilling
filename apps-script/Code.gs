/**
 * UpSkilling - Backend
 * Fonte oficial: Google Apps Script
 * Esta versão reflete o estado em produção
 */

const SHEET_SKILLS = "Skills";
const SHEET_LOGS = "Logs";

function doGet(e) {
  const action = e.parameter.action;

  if (action === "getSkills") return getSkills();
  if (action === "getLogs") return getLogs();
  if (action === "getQuests") return getQuests();

  return json({ error: "Invalid action" });
}

function getLogs() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Logs");
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const logs = data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  return json(logs);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if (action === "addXP") {
    return addXP(data);
  }

  if (action === "createSkill") {
    return createSkill(data);
  }

  if (action === "completeQuest") {
  return completeQuest(data);
}

  return ContentService.createTextOutput(
    JSON.stringify({ error: "Invalid action" })
  ).setMimeType(ContentService.MimeType.JSON);
}

// =====================
// SKILLS
// =====================

function getSkills() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SKILLS);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const skills = data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });

  return json(skills);
}

function createSkill(data) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SKILLS);
  const lastRow = sheet.getLastRow();

  const newId = lastRow;
  sheet.appendRow([
    newId,
    data.name,
    data.description,
    new Date()
  ]);

  return json({ success: true });
}

// =====================
// XP / LOGS
// =====================

function addXP(data) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);

  sheet.appendRow([
    new Date(),
    data.skill_id,
    data.points,
    data.source || "manual"
  ]);

  return json({ success: true });
}

// =====================
// HELPERS
// =====================

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// =====================
// QUESTS
// =====================

function getQuests() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Quests");
  const data = sheet.getDataRange().getValues();

  const headers = data.shift();
  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const quests = data.map(row => ({
    id: row[idx.id],
    name: row[idx.name],
    description: row[idx.description],
    reward_type: row[idx.reward_type],
    value: row[idx.value],
    multiplier: row[idx.multiplier],
    duration_days: row[idx.duration_days],
    skill_target: row[idx.skill_target],
    done: row[idx.done]
  }));

  return json(quests);
}

function completeQuest(data) {
  const questId = data.quest_id;

  const ss = SpreadsheetApp.getActive();
  const questSheet = ss.getSheetByName("Quests");
  const logSheet = ss.getSheetByName("Logs");
  const metaSheet = ss.getSheetByName("Meta");

  const values = questSheet.getDataRange().getValues();
  const headers = values.shift();

  const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

  const rowIndex = values.findIndex(r => r[idx.id] == questId);
  if (rowIndex === -1) {
    return json({ success: false, error: "Quest não encontrada" });
  }

  const row = values[rowIndex];
  const sheetRow = rowIndex + 2;

  // 🔒 Se já foi concluída, não faz nada
  if (row[idx.done]) {
    return json({ success: false, error: "Quest já concluída" });
  }

  // ✅ Marca como concluída
  questSheet
    .getRange(sheetRow, idx.done + 1)
    .setValue(new Date());

  const rewardType = row[idx.reward_type];

  // 🎯 RECOMPENSA XP
  if (rewardType === "xp") {
    const xpValues = row[idx.value]
      .toString()
      .split(",")
      .map(v => Number(v.trim()));

    const skillTargets = row[idx.skill_target]
      .toString()
      .split(",")
      .map(v => v.trim());

    const now = new Date();

    xpValues.forEach((xp, i) => {
      logSheet.appendRow([
        now,
        skillTargets[i],
        xp,
        "Quest"
      ]);
    });
  }

  // ✨ RECOMPENSA MULTIPLICADOR
  if (rewardType === "mult") {
    const multiplier = row[idx.multiplier];
    const days = row[idx.duration_days];

    const expiresAt = new Date();
    expiresAt.setDate(expiresAt.getDate() + Number(days));

    metaSheet.appendRow([
      "multiplier",
      multiplier,
      new Date(),
      expiresAt
    ]);
  }

  return json({ success: true });
}

