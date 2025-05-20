/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import LZString from "lz-string";
import { ONEDRIVE_CLIENT_ID } from "./secrets";

// OneDrive obsługa
let useOneDrive = false;
let graphAccessToken: string | null = null;
const ONEDRIVE_FOLDER = "OutlookNotesAddIn";
const ONEDRIVE_FILE_PREFIX = "note_";

// Importuj MSAL tylko jeśli jest potrzebny
let msalInstance: any = null;
if (ONEDRIVE_CLIENT_ID && (ONEDRIVE_CLIENT_ID as string) !== "YOUR_ONEDRIVE_CLIENT_ID_HERE") {
  // Dynamiczny import, żeby nie ładować MSAL bez potrzeby
  import("@azure/msal-browser").then(msal => {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: ONEDRIVE_CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin
      }
    });
  });
}

async function tryEnableOneDrive() {
  if (!msalInstance) return false;
  try {
    const loginResponse = await msalInstance.loginPopup({
      scopes: ["Files.ReadWrite", "User.Read"]
    });
    graphAccessToken = loginResponse.accessToken;
    useOneDrive = true;
    await ensureOneDriveFolder();
    return true;
  } catch {
    useOneDrive = false;
    return false;
  }
}

async function ensureOneDriveFolder() {
  if (!graphAccessToken) return;
  const res = await fetch("https://graph.microsoft.com/v1.0/me/drive/special/approot/children", {
    headers: { Authorization: `Bearer ${graphAccessToken}` }
  });
  const data = await res.json();
  if (!data.value.some((f: any) => f.name === ONEDRIVE_FOLDER)) {
    await fetch("https://graph.microsoft.com/v1.0/me/drive/special/approot/children", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${graphAccessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ name: ONEDRIVE_FOLDER, folder: {}, "@microsoft.graph.conflictBehavior": "rename" })
    });
  }
}

async function saveToOneDrive(noteData: NoteData) {
  if (!graphAccessToken || !currentEntryId) return;
  const fileName = `${ONEDRIVE_FILE_PREFIX}${currentEntryId}.json`;
  await fetch(`https://graph.microsoft.com/v1.0/me/drive/special/approot:/${ONEDRIVE_FOLDER}/${fileName}:/content`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${graphAccessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(noteData)
  });
}

async function loadFromOneDrive(): Promise<NoteData | null> {
  if (!graphAccessToken || !currentEntryId) return null;
  const fileName = `${ONEDRIVE_FILE_PREFIX}${currentEntryId}.json`;
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/special/approot:/${ONEDRIVE_FOLDER}/${fileName}:/content`, {
    headers: { Authorization: `Bearer ${graphAccessToken}` }
  });
  if (res.ok) {
    return await res.json();
  }
  return null;
}

// --- Główna logika aplikacji ---

interface TodoItem {
  text: string;
  isDone: boolean;
}

interface NoteData {
  text: string;
  todos: TodoItem[];
}

const NOTES_KEY = "notesData";
let currentEntryId: string | null = null;
let loadedText = "";
let todos: TodoItem[] = [];
let uiInitialized = false;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Jeśli jest prawidłowy clientId, zapytaj o OneDrive
    if (ONEDRIVE_CLIENT_ID && (ONEDRIVE_CLIENT_ID as string) !== "YOUR_ONEDRIVE_CLIENT_ID_HERE" && msalInstance) {
      const agree = confirm("Czy chcesz przechowywać notatki na swoim OneDrive?");
      if (agree) {
        await tryEnableOneDrive();
      }
    }
    runOutlook();
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => {
        saveCurrent();
        runOutlook();
      }
    );
  }

  const loaded = await loadNote();
  if (loaded) {
    (document.getElementById("txtNote") as HTMLTextAreaElement).value = loaded.text;
    todos = loaded.todos;
    refreshList();
  }

  window.addEventListener("beforeunload", () => {
    saveCurrent();
  });
});

export async function runOutlook() {
  const item = Office.context.mailbox.item;
  const newEntryId = item.conversationId || item.itemId || item.internetMessageId || "unknown";
  currentEntryId = newEntryId;
  setupUI();
  await loadNote();
}

function setupUI() {
  if (uiInitialized) return;
  document.getElementById("txtNote").addEventListener("input", txtNote_Changed);
  document.getElementById("btnSave").addEventListener("click", btnSave_Click);
  document.getElementById("txtNewTodo").addEventListener("input", txtNewTodo_Changed);
  document.getElementById("btnAddTodo").addEventListener("click", btnAddTodo_Click);

  document.getElementById("txtNewTodo").addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter" && !(e.shiftKey || e.ctrlKey || e.altKey)) {
      e.preventDefault();
      if (!(document.getElementById("btnAddTodo") as HTMLButtonElement).disabled) {
        btnAddTodo_Click();
      }
    }
  });

  txtNewTodo_Changed();
  uiInitialized = true;
}

function txtNote_Changed() {
  const txtNote = (document.getElementById("txtNote") as HTMLTextAreaElement).value;
  document.getElementById("btnSave").toggleAttribute("disabled", txtNote === loadedText);
}

function txtNewTodo_Changed() {
  const txt = (document.getElementById("txtNewTodo") as HTMLInputElement).value.trim();
  const btn = document.getElementById("btnAddTodo") as HTMLButtonElement;
  btn.disabled = txt.length === 0;
}

function btnSave_Click() {
  const noteData: NoteData = {
    text: (document.getElementById("txtNote") as HTMLTextAreaElement).value,
    todos: todos
  };
  saveToRoamingSettings(noteData);
  saveCurrent();
}

function btnAddTodo_Click() {
  const txtInput = document.getElementById("txtNewTodo") as HTMLInputElement;
  const text = txtInput.value.trim();
  if (!text) return;
  todos.push({ text, isDone: false });
  txtInput.value = "";
  refreshList();
  txtNewTodo_Changed();
}

function refreshList() {
  const ul = document.getElementById("todos");
  ul.innerHTML = "";
  todos.forEach((td, idx) => {
    const li = document.createElement("li");
    li.innerHTML = `
      <div class="todo-row" draggable="true" data-idx="${idx}">
        <div class="todo-handle-col" title="Przeciągnij, aby zmienić kolejność">
          <span class="todo-handle" aria-label="Przeciągnij">&#9776;</span>
        </div>
        <div class="todo-check-col">
          <input type="checkbox" ${td.isDone ? "checked" : ""} data-idx="${idx}" />
        </div>
        <div class="todo-text-col">
          <span >${td.text}</span>
        </div>
        <div class="todo-delete-col">
          <span class="delete-todo" data-del="${idx}" title="Usuń">
            <svg class="fluent-delete-icon" width="20" height="20" viewBox="0 0 20 20" fill="none">
              <path d="M8.5 4H11.5C11.5 3.17157 10.8284 2.5 10 2.5C9.17157 2.5 8.5 3.17157 8.5 4ZM7.5 4C7.5 2.61929 8.61929 1.5 10 1.5C11.3807 1.5 12.5 2.61929 12.5 4H17.5C17.7761 4 18 4.22386 18 4.5C18 4.77614 17.7761 5 17.5 5H16.4456L15.2521 15.3439C15.0774 16.8576 13.7957 18 12.2719 18H7.72813C6.20431 18 4.92256 16.8576 4.7479 15.3439L3.55437 5H2.5C2.22386 5 2 4.77614 2 4.5C2 4.22386 2.22386 4 2.5 4H7.5ZM5.74131 15.2292C5.85775 16.2384 6.71225 17 7.72813 17H12.2719C13.2878 17 14.1422 16.2384 14.2587 15.2292L15.439 5H4.56101L5.74131 15.2292ZM8.5 7.5C8.77614 7.5 9 7.72386 9 8V14C9 14.2761 8.77614 14.5 8.5 14.5C8.22386 14.5 8 14.2761 8 14V8C8 7.72386 8.22386 7.5 8.5 7.5ZM12 8C12 7.72386 11.7761 7.5 11.5 7.5C11.2239 7.5 11 7.72386 11 8V14C11 14.2761 11.2239 14.5 11.5 14.5C11.7761 14.5 12 14.2761 12 14V8Z" fill="currentColor"/>
            </svg>
          </span>
        </div>
      </div>
    `;
    ul.appendChild(li);
  });

  const rows = ul.querySelectorAll(".todo-row");
  let dragSrcIdx: number | null = null;

  rows.forEach(row => {
    row.addEventListener("dragstart", (e) => {
      dragSrcIdx = +(row as HTMLElement).dataset.idx;
      row.classList.add("dragging");
      (e as DragEvent).dataTransfer.effectAllowed = "move";
    });
    row.addEventListener("dragend", () => {
      row.classList.remove("dragging");
      dragSrcIdx = null;
      rows.forEach(r => r.classList.remove("drag-over"));
    });
    row.addEventListener("dragover", (e) => {
      e.preventDefault();
      row.classList.add("drag-over");
    });
    row.addEventListener("dragleave", () => {
      row.classList.remove("drag-over");
    });
    row.addEventListener("drop", (e) => {
      e.preventDefault();
      row.classList.remove("drag-over");
      const dropIdx = +(row as HTMLElement).dataset.idx;
      if (dragSrcIdx !== null && dragSrcIdx !== dropIdx) {
        const moved = todos.splice(dragSrcIdx, 1)[0];
        todos.splice(dropIdx, 0, moved);
        refreshList();
        document.getElementById("btnSave").removeAttribute("disabled");
      }
    });
  });

  ul.querySelectorAll("input[type=checkbox]").forEach(cb => {
    cb.addEventListener("change", (e) => {
      const idx = +(e.target as HTMLInputElement).dataset.idx;
      todos[idx].isDone = (e.target as HTMLInputElement).checked;
      saveCurrent();
      refreshList();
    });
  });
  ul.querySelectorAll(".delete-todo").forEach(btn => {
    btn.addEventListener("click", (e) => {
      const idx = +(e.currentTarget as HTMLElement).dataset.del;
      if (idx >= 0 && idx < todos.length) {
        todos.splice(idx, 1);
        refreshList();
        document.getElementById("btnSave").removeAttribute("disabled");
      }
    });
  });
}

async function saveCurrent() {
  if (!currentEntryId) return;
  const txtNote = (document.getElementById("txtNote") as HTMLTextAreaElement).value;
  const noteData: NoteData = { text: txtNote, todos };
  if (useOneDrive) {
    await saveToOneDrive(noteData);
  } else {
    saveToRoamingSettings(noteData);
  }
  loadedText = txtNote;
  document.getElementById("btnSave").setAttribute("disabled", "true");
}

async function loadNote() {
  if (!currentEntryId) return;
  let noteData: NoteData = { text: "", todos: [] };
  if (useOneDrive) {
    const loaded = await loadFromOneDrive();
    if (loaded) noteData = loaded;
  } else {
    const loaded = loadFromRoamingSettings();
    if (loaded) noteData = loaded;
  }
  loadedText = noteData.text;
  todos = noteData.todos || [];
  (document.getElementById("txtNote") as HTMLTextAreaElement).value = loadedText;
  refreshList();
  document.getElementById("btnSave").setAttribute("disabled", "true");
  return noteData;
}

function saveToRoamingSettings(noteData: NoteData) {
  if (!currentEntryId) return;
  // Kompresuj dane przed zapisem
  const compressed = LZString.compressToUTF16(JSON.stringify(noteData));
  Office.context.roamingSettings.set(NOTES_KEY + "_" + currentEntryId, compressed);
  Office.context.roamingSettings.saveAsync();
}

function loadFromRoamingSettings(): NoteData | null {
  if (!currentEntryId) return null;
  const data = Office.context.roamingSettings.get(NOTES_KEY + "_" + currentEntryId);
  if (data) {
    try {
      // Dekompresuj dane po odczycie
      return JSON.parse(LZString.decompressFromUTF16(data));
    } catch {
      return null;
    }
  }
  return null;
}







