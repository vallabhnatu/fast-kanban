// --- CONFIGURATION ---
const DB_SCHEMA = {
  Tickets: [
    "id",
    "title",
    "priority",
    "status",
    "description",
    "assignee",
    "sprintId",
    "due",
    "created",
    "updated",
  ],
  Sprints: ["id", "name", "status", "startDate", "endDate", "completedDate"],
  Settings: ["key", "value"],
  Columns: ["id", "title", "orderIndex"],
};

// --- INITIALIZATION ---

/**
 * RUN THIS FUNCTION FIRST to set up the spreadsheet structure.
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(DB_SCHEMA).forEach((sheetName) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(DB_SCHEMA[sheetName]); // Add Headers

      // Add default data if creating for the first time
      if (sheetName === "Settings") {
        sheet.appendRow(["projectKey", "KAN"]);
        sheet.appendRow(["ticketCounter", "100"]);
      }
      if (sheetName === "Columns") {
        sheet.appendRow(["todo", "To Do", "0"]);
        sheet.appendRow(["progress", "In Progress", "1"]);
        sheet.appendRow(["done", "Done", "2"]);
      }
      if (sheetName === "Sprints") {
        const now = new Date().toISOString();
        // Create initial active sprint
        sheet.appendRow(["s1", "Sprint 1", "active", now, "", ""]);
      }
    }
  });
}

// --- API HANDLERS ---

function doGet(e) {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("FastKanban")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no"
    );
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  // Wait for up to 10 seconds for other processes to finish
  if (!lock.tryLock(10000)) {
    return response({
      status: "error",
      message: "Server is busy. Please try again.",
    });
  }

  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;
    let result;

    switch (action) {
      case "loadInitialData":
        result = loadInitialData();
        break;
      case "createTicket":
        result = createTicket(data);
        break;
      case "updateTicket":
        result = updateTicket(data);
        break;
      case "deleteTicket":
        result = deleteTicket(data);
        break;
      case "createColumn":
        result = createColumn(data);
        break;
      case "deleteColumn":
        result = deleteColumn(data);
        break;
      case "reorderColumns":
        result = reorderColumns(data);
        break;
      case "completeSprint":
        result = completeSprint();
        break;
      case "updateSettings":
        result = updateSettings(data);
        break;
      default:
        throw new Error("Unknown action: " + action);
    }

    return response({ status: "success", data: result });
  } catch (err) {
    return response({
      status: "error",
      message: err.toString(),
      stack: err.stack,
    });
  } finally {
    lock.releaseLock();
  }
}

// --- CORE LOGIC ---

function loadInitialData() {
  const tickets = getSheetData("Tickets");
  const sprints = getSheetData("Sprints");
  const columns = getSheetData("Columns").sort(
    (a, b) => a.orderIndex - b.orderIndex
  );
  const settingsRaw = getSheetData("Settings");

  // Convert settings array to object
  const settings = settingsRaw.reduce((acc, row) => {
    acc[row.key] = row.value;
    return acc;
  }, {});

  return { tickets, sprints, columns, settings };
}

function createTicket(ticketData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const data = settingsSheet.getDataRange().getValues();

  // 1. Get and Increment Counter
  let counterRowIndex = data.findIndex((r) => r[0] === "ticketCounter");
  let projectKeyRow = data.find((r) => r[0] === "projectKey");

  let counter = parseInt(data[counterRowIndex][1]);
  let projectKey = projectKeyRow ? projectKeyRow[1] : "KAN";

  counter++;
  settingsSheet.getRange(counterRowIndex + 1, 2).setValue(counter);

  // 2. Construct ID and Object
  const newId = `${projectKey}-${counter}`;
  const now = new Date().toISOString();

  const newTicket = {
    id: newId,
    title: ticketData.title,
    priority: ticketData.priority || "Medium",
    status: ticketData.status || "backlog",
    description: ticketData.description || "",
    assignee: ticketData.assignee || "",
    sprintId: ticketData.sprintId || "",
    due: ticketData.due || "",
    created: now,
    updated: now,
  };

  // 3. Save to Sheet
  addRow("Tickets", newTicket);
  return newTicket;
}

function updateTicket(updateData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Tickets");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find row by ID (ID is column 0)
  const rowIndex = data.findIndex((r) => r[0] === updateData.id);

  if (rowIndex === -1) throw new Error("Ticket not found");

  // Map updates to column indices
  const rowNumber = rowIndex + 1;
  const currentRow = data[rowIndex];

  updateData.updated = new Date().toISOString(); // Always update timestamp

  Object.keys(updateData).forEach((key) => {
    const colIndex = headers.indexOf(key);
    if (colIndex !== -1) {
      // Update specific cell
      sheet.getRange(rowNumber, colIndex + 1).setValue(updateData[key]);
      // Update memory for return
      currentRow[colIndex] = updateData[key];
    }
  });

  // Return updated object
  return mapRowToObject(currentRow, headers);
}

function deleteTicket(data) {
  deleteRowById("Tickets", data.id);
  return { id: data.id };
}

function createColumn(data) {
  const columns = getSheetData("Columns");
  const maxOrder = columns.reduce(
    (max, c) => Math.max(max, parseInt(c.orderIndex || 0)),
    0
  );

  const newCol = {
    id: data.title.toLowerCase().replace(/[^a-z0-9]/g, "-"),
    title: data.title,
    orderIndex: maxOrder + 1,
  };

  // Check duplicate
  if (columns.find((c) => c.id === newCol.id)) return newCol;

  addRow("Columns", newCol);
  return newCol;
}

function deleteColumn(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ticketSheet = ss.getSheetByName("Tickets");
  const tickets = getSheetData("Tickets");

  // 1. Move tickets in this column to backlog
  tickets.forEach((t, index) => {
    if (t.status === data.id) {
      // Row index is index + 2 (1 for 0-base, 1 for header)
      // Update Status (col index 3 in schema, 4 in 1-base)
      // Ideally we dynamic lookup
      const headers = DB_SCHEMA.Tickets;
      const statusIdx = headers.indexOf("status") + 1;
      const sprintIdx = headers.indexOf("sprintId") + 1;
      const updateIdx = headers.indexOf("updated") + 1;

      ticketSheet.getRange(index + 2, statusIdx).setValue("backlog");
      ticketSheet.getRange(index + 2, sprintIdx).setValue("");
      ticketSheet
        .getRange(index + 2, updateIdx)
        .setValue(new Date().toISOString());
    }
  });

  // 2. Delete Column Definition
  deleteRowById("Columns", data.id);
  return { success: true };
}

function reorderColumns(data) {
  // data.newOrderIds = ['todo', 'qa', 'done']
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Columns");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idIdx = headers.indexOf("id");
  const orderIdx = headers.indexOf("orderIndex");

  for (let i = 1; i < rows.length; i++) {
    const id = rows[i][idIdx];
    const newIndex = data.newOrderIds.indexOf(id);
    if (newIndex !== -1) {
      sheet.getRange(i + 1, orderIdx + 1).setValue(newIndex);
    }
  }
  return { success: true };
}

function completeSprint() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get Sprints
  const sprintSheet = ss.getSheetByName("Sprints");
  const sprintsRaw = sprintSheet.getDataRange().getValues();
  const sprintHeaders = sprintsRaw[0];
  const statusIdx = sprintHeaders.indexOf("status");
  const idIdx = sprintHeaders.indexOf("id");
  const completedDateIdx = sprintHeaders.indexOf("completedDate");

  // Find active sprint
  let activeRowIndex = -1;
  let activeSprintId = null;

  for (let i = 1; i < sprintsRaw.length; i++) {
    if (sprintsRaw[i][statusIdx] === "active") {
      activeRowIndex = i + 1;
      activeSprintId = sprintsRaw[i][idIdx];
      break;
    }
  }

  if (!activeSprintId) throw new Error("No active sprint found");

  const now = new Date().toISOString();

  // 2. Close Active Sprint
  sprintSheet.getRange(activeRowIndex, statusIdx + 1).setValue("completed");
  sprintSheet.getRange(activeRowIndex, completedDateIdx + 1).setValue(now);

  // 3. Create New Sprint
  const sprintCount = sprintsRaw.length; // Header + data + 1 for new
  const newSprintId = "s" + sprintCount;
  const newSprint = {
    id: newSprintId,
    name: "Sprint " + sprintCount,
    status: "active",
    startDate: now,
    endDate: "",
    completedDate: "",
  };
  addRow("Sprints", newSprint);

  // 4. Move incomplete tickets
  const ticketSheet = ss.getSheetByName("Tickets");
  const tickets = getSheetData("Tickets");
  const tHeaders = DB_SCHEMA.Tickets;
  const tStatusIdx = tHeaders.indexOf("status");
  const tSprintIdx = tHeaders.indexOf("sprintId");

  tickets.forEach((t, i) => {
    if (t.sprintId === activeSprintId && t.status !== "done") {
      // Move to new sprint
      ticketSheet.getRange(i + 2, tSprintIdx + 1).setValue(newSprintId);
    }
  });

  return {
    completedSprintId: activeSprintId,
    newSprint: newSprint,
  };
}

function updateSettings(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  const rows = sheet.getDataRange().getValues();

  if (data.projectKey) {
    const oldKeyRow = rows.find((r) => r[0] === "projectKey");
    const oldKey = oldKeyRow ? oldKeyRow[1] : "KAN";
    const newKey = data.projectKey;

    if (oldKey !== newKey) {
      // Update Setting
      const keyRowIndex = rows.findIndex((r) => r[0] === "projectKey");
      sheet.getRange(keyRowIndex + 1, 2).setValue(newKey);

      // Update ALL Ticket IDs
      const ticketSheet = ss.getSheetByName("Tickets");
      const tData = ticketSheet.getDataRange().getValues();
      const tIdIdx = 0; // ID is always first

      // Batch update for speed
      const updates = [];
      for (let i = 1; i < tData.length; i++) {
        const currentId = tData[i][tIdIdx];
        if (currentId.startsWith(oldKey + "-")) {
          const newId = currentId.replace(oldKey, newKey);
          ticketSheet.getRange(i + 1, 1).setValue(newId);
        }
      }
    }
  }

  return { success: true };
}

// --- HELPERS ---

function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return []; // Only headers

  const headers = data[0];
  return data.slice(1).map((row) => mapRowToObject(row, headers));
}

function mapRowToObject(row, headers) {
  const obj = {};
  headers.forEach((h, i) => {
    obj[h] = row[i];
  });
  return obj;
}

function addRow(sheetName, dataObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const headers = DB_SCHEMA[sheetName];

  const row = headers.map((h) => {
    const val = dataObj[h];
    return val === undefined || val === null ? "" : val;
  });

  sheet.appendRow(row);
}

function deleteRowById(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();

  // Assuming ID is always first column
  const rowIndex = data.findIndex((r) => r[0] === id);
  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
  }
}

function response(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
