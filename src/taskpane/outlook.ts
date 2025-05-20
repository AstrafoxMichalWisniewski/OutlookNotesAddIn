/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import LZString from "lz-string";


interface TodoItem {
  text: string;
  isDone: boolean;
}
function makeKey(entryId: string): string {
  return `OutlookNotesAddIn_${entryId}`;
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


async function saveToCustomProps(noteData: NoteData, entryId: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const compressed = LZString.compressToUTF16(JSON.stringify(noteData));
    Office.context.mailbox.item.loadCustomPropertiesAsync(loadResult => {
      if (loadResult.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(loadResult.error);
      }
      const props = loadResult.value;
      props.set(makeKey(entryId), compressed);
      props.saveAsync(saveResult => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(saveResult.error);
        }
      });
    });
  });
}

async function loadFromCustomProps(entryId: string): Promise<NoteData | null> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.loadCustomPropertiesAsync(loadResult => {
      if (loadResult.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(loadResult.error);
      }
      const props = loadResult.value;
      const data = props.get(makeKey(entryId)) as string | undefined;
      if (!data) {
        return resolve(null);
      }
      try {
        const obj = JSON.parse(LZString.decompressFromUTF16(data) || "{}") as NoteData;
        resolve(obj);
      } catch {
        resolve(null);
      }
    });
  });
}
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    await runOutlook();
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      async () => {
        await saveCurrent();
        await runOutlook();
      }
    );
  }
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
    const txtNoteElem = document.getElementById("txtNote") as HTMLTextAreaElement;

  txtNoteElem.addEventListener("input", txtNote_Changed);
  txtNoteElem.addEventListener("blur", async () => {
    if (txtNoteElem.value !== loadedText) {
      await saveCurrent();
    }
  });



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

async function btnSave_Click() {
  await saveCurrent();
}

async function btnAddTodo_Click() {
  const txtInput = document.getElementById("txtNewTodo") as HTMLInputElement;
  const text = txtInput.value.trim();
  if (!text) return;

  todos.push({ text, isDone: false });
  txtInput.value = "";
  refreshList();
  txtNewTodo_Changed();
  await saveCurrent(); // << dopisane
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
  cb.addEventListener("change", async (e) => {
    const idx = +(e.target as HTMLInputElement).dataset.idx;
    todos[idx].isDone = (e.target as HTMLInputElement).checked;
    await saveCurrent();
  });
});
ul.querySelectorAll(".delete-todo").forEach(btn => {
  btn.addEventListener("click", async (e) => {
    const idx = +(e.currentTarget as HTMLElement).dataset.del;
    if (idx >= 0 && idx < todos.length) {
      todos.splice(idx, 1);
      refreshList();
      await saveCurrent();
    }
  });
});
}

async function loadNote() {
  if (!currentEntryId) return null;

  const noteData = await loadFromCustomProps(currentEntryId);
  if (!noteData) {
    loadedText = "";
    todos = [];
    (document.getElementById("txtNote") as HTMLTextAreaElement).value = "";
    refreshList();
    (document.getElementById("btnSave") as HTMLButtonElement).disabled = true;
    return null;
  }

  loadedText = noteData.text;
  todos = noteData.todos || [];

  (document.getElementById("txtNote") as HTMLTextAreaElement).value = loadedText;
  refreshList();
  (document.getElementById("btnSave") as HTMLButtonElement).disabled = true;

  return noteData;
}

async function saveCurrent(entryIdOverride?: string) {
  const entryIdToUse = entryIdOverride || currentEntryId;
  if (!entryIdToUse) return;

  const txtNote = (document.getElementById("txtNote") as HTMLTextAreaElement).value;
  const noteData: NoteData = { text: txtNote, todos };
  await saveToCustomProps(noteData, entryIdToUse);
  if (!entryIdOverride) {
    loadedText = txtNote;
    (document.getElementById("btnSave") as HTMLButtonElement).disabled = true;
  }
}







