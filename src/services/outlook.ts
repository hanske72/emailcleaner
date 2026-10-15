import { getGraphToken, getGraphTokenSafe } from "./graphAuth";

let _useStrictTokenMode = false;

/** When enabled, Graph calls throw instead of opening auth dialog on token expiry */
export function setStrictTokenMode(on: boolean): void {
  _useStrictTokenMode = on;
}

export interface MailFolder {
  id: string;
  name: string;
  totalCount: number;
  unreadCount: number;
}

export interface MailMessage {
  id: string;
  index: number;
  subject: string;
  fromName: string;
  fromAddress: string;
  date: string;
  isRead: boolean;
  folderId: string;
  folderName: string;
}

export interface FolderSnapshot {
  id: string;
  name: string;
  path: string;            // full path e.g. "Innboks/_1_Kunder/Fremtind"
  totalCount: number;
  unreadCount: number;
  messages: MailMessage[];  // inbox = all messages, user folders = sample (up to 5)
  children: FolderSnapshot[];
}

export interface MailboxSnapshot {
  folders: FolderSnapshot[];
  totalMessages: number;
  totalUnread: number;
  allMessages: MailMessage[];
}

const GRAPH = "https://graph.microsoft.com/v1.0/me";

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

const GRAPH_TIMEOUT_MS = 15000;

async function graphFetch(path: string, init?: RequestInit, retries = 3): Promise<Response> {
  const token = _useStrictTokenMode ? await getGraphTokenSafe() : await getGraphToken();
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), GRAPH_TIMEOUT_MS);
  let res: Response;
  try {
    res = await fetch(`${GRAPH}${path}`, {
      ...init,
      signal: controller.signal,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(init?.headers || {}),
      },
    });
  } catch (e: any) {
    clearTimeout(timer);
    if (e.name === "AbortError") throw new Error(`Graph timeout etter ${GRAPH_TIMEOUT_MS / 1000}s: ${path}`);
    throw e;
  }
  clearTimeout(timer);
  if (res.status === 429 && retries > 0) {
    const retryAfter = parseInt(res.headers.get("Retry-After") || "5", 10);
    await delay(retryAfter * 1000);
    return graphFetch(path, init, retries - 1);
  }
  if (res.status === 401) {
    throw new Error("Tilgang nektet. Prøv å logge inn på nytt.");
  }
  if (!res.ok && init?.method !== "DELETE") {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph API-feil ${res.status}: ${text || res.statusText}`);
  }
  return res;
}

async function graphGet(path: string): Promise<any> {
  const res = await graphFetch(path);
  return res.json();
}

async function graphPost(path: string, body: any): Promise<any> {
  const res = await graphFetch(path, {
    method: "POST",
    body: JSON.stringify(body),
  });
  return res.json();
}

const INBOX_NAMES = new Set(["innboks", "inbox"]);
const SYSTEM_FOLDER_NAMES = new Set([
  "innboks", "inbox",
  "sendte elementer", "sent items",
  "kladd", "drafts",
  "utboks", "outbox",
  "slettede elementer", "deleted items",
  "søppelpost", "junk email", "junk", "spam",
  "arkiv", "archive",
  "samtalelogg", "conversation history",
]);

let _globalIndex = 0;

/** Recursively scan a folder and its children */
async function scanFolderTree(
  folderId: string,
  folderName: string,
  parentPath: string,
  totalCount: number,
  unreadCount: number,
  isTopLevelInbox: boolean,
  onProgress: (msg: string, pct: number) => void,
  pct: number,
): Promise<FolderSnapshot> {
  const path = parentPath ? `${parentPath}/${folderName}` : folderName;
  const isSystem = SYSTEM_FOLDER_NAMES.has(folderName.toLowerCase());

  let messages: MailMessage[] = [];

  if (totalCount > 0) {
    if (isTopLevelInbox) {
      // Inbox: fetch up to 500 messages
      onProgress(`Skanner ${path} (${totalCount} e-poster)...`, pct);
      try {
        let nextLink: string | null = `/mailFolders/${folderId}/messages?$top=100&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead`;
        while (nextLink && messages.length < 500) {
          const msgRes = await graphGet(nextLink);
          const batch = ((msgRes.value || []) as any[]).map((m: any) => ({
            id: m.id,
            index: _globalIndex++,
            subject: m.subject || "(uten emne)",
            fromName: m.from?.emailAddress?.name || "",
            fromAddress: m.from?.emailAddress?.address || "",
            date: m.receivedDateTime || "",
            isRead: m.isRead ?? true,
            folderId: folderId,
            folderName: folderName,
          }));
          messages.push(...batch);
          const raw = msgRes["@odata.nextLink"];
          nextLink = raw ? raw.replace("https://graph.microsoft.com/v1.0/me", "") : null;
        }
      } catch { /* skip unreadable */ }
    } else if (!isSystem) {
      // User folder: fetch small sample (5 newest) to understand content
      try {
        const msgRes = await graphGet(
          `/mailFolders/${folderId}/messages?$top=5&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead`
        );
        messages = ((msgRes.value || []) as any[]).map((m: any) => ({
          id: m.id,
          index: -1, // sample messages, not actionable
          subject: m.subject || "(uten emne)",
          fromName: m.from?.emailAddress?.name || "",
          fromAddress: m.from?.emailAddress?.address || "",
          date: m.receivedDateTime || "",
          isRead: m.isRead ?? true,
          folderId: folderId,
          folderName: folderName,
        }));
      } catch { /* skip */ }
    }
  }

  // Scan child folders (subfolders)
  let children: FolderSnapshot[] = [];
  try {
    const childRes = await graphGet(`/mailFolders/${folderId}/childFolders?$top=100`);
    const childFolders: any[] = childRes.value || [];
    for (const cf of childFolders) {
      const child = await scanFolderTree(
        cf.id,
        cf.displayName || "Ukjent",
        path,
        cf.totalItemCount || 0,
        cf.unreadItemCount || 0,
        false, // subfolders are never the inbox itself
        onProgress,
        pct,
      );
      children.push(child);
    }
  } catch { /* no children or permission denied */ }

  return {
    id: folderId,
    name: folderName,
    path,
    totalCount,
    unreadCount,
    messages,
    children,
  };
}

export async function scanMailbox(
  onProgress: (msg: string, pct: number) => void
): Promise<MailboxSnapshot> {
  onProgress("Henter mapper...", 5);
  const foldersRes = await graphGet("/mailFolders?$top=100");
  const rawFolders: any[] = foldersRes.value || [];

  _globalIndex = 0;

  const snapshot: MailboxSnapshot = {
    folders: [],
    totalMessages: 0,
    totalUnread: 0,
    allMessages: [],
  };

  for (let i = 0; i < rawFolders.length; i++) {
    const f = rawFolders[i];
    const name = f.displayName || "Ukjent";
    const total = f.totalItemCount || 0;
    const unread = f.unreadItemCount || 0;
    const pct = 5 + Math.round(((i + 1) / rawFolders.length) * 90);
    const isInbox = INBOX_NAMES.has(name.toLowerCase());

    if (total > 0 || isInbox) {
      onProgress(`Skanner ${name}...`, pct);
    }

    const folderSnap = await scanFolderTree(
      f.id, name, "", total, unread, isInbox, onProgress, pct
    );

    snapshot.folders.push(folderSnap);
    snapshot.totalMessages += total;
    snapshot.totalUnread += unread;

    // Collect inbox messages for allMessages
    if (isInbox) {
      snapshot.allMessages.push(...folderSnap.messages);
    }
  }

  onProgress("Skanning fullført!", 100);
  return snapshot;
}

export async function moveMessages(ids: string[], targetFolderId: string): Promise<number> {
  let ok = 0;
  for (const id of ids) {
    try {
      await graphPost(`/messages/${id}/move`, { destinationId: targetFolderId });
      ok++;
      if (ids.length > 3) await delay(200); // throttle to avoid 429
    } catch { /* continue */ }
  }
  return ok;
}

export async function createFolder(name: string, parentId?: string): Promise<string> {
  const endpoint = parentId
    ? `/mailFolders/${parentId}/childFolders`
    : "/mailFolders";
  try {
    const data = await graphPost(endpoint, { displayName: name });
    return data.id;
  } catch (e: any) {
    // Folder already exists (409) — find it by name
    if (e.message?.includes("409") || e.message?.includes("ErrorFolderExists")) {
      const searchEndpoint = parentId
        ? `/mailFolders/${parentId}/childFolders?$filter=displayName eq '${name.replace(/'/g, "''")}'`
        : `/mailFolders?$filter=displayName eq '${name.replace(/'/g, "''")}'`;
      const res = await graphGet(searchEndpoint);
      if (res.value?.length > 0) return res.value[0].id;
    }
    throw e;
  }
}

export async function deleteMessages(ids: string[]): Promise<number> {
  let ok = 0;
  for (const id of ids) {
    try {
      const res = await graphFetch(`/messages/${id}`, { method: "DELETE" });
      if (res.ok) ok++;
      if (ids.length > 3) await delay(200); // throttle to avoid 429
    } catch { /* continue */ }
  }
  return ok;
}

/** Names for the "Other" focused-inbox folder in various locales */
const OTHER_INBOX_NAMES = new Set(["other", "annet", "andere", "autre", "otro"]);

/** Scan folder structure WITHOUT fetching inbox messages (for v4.0 auto-clean) */
export async function scanFolderStructure(
  onProgress: (msg: string, pct: number) => void
): Promise<{ folders: FolderSnapshot[]; inboxId: string; inboxCount: number; otherInboxId?: string; otherInboxCount: number }> {
  onProgress("Henter mapper...", 5);

  // Single API call with $expand to get top-level folders + their children in one request
  const foldersRes = await graphGet(
    "/mailFolders?$top=100&$expand=childFolders($top=100)"
  );
  const rawFolders: any[] = foldersRes.value || [];

  let inboxId = "";
  let inboxCount = 0;
  let otherInboxId: string | undefined;
  let otherInboxCount = 0;
  const folders: FolderSnapshot[] = [];

  for (let i = 0; i < rawFolders.length; i++) {
    const f = rawFolders[i];
    const name = f.displayName || "Ukjent";
    const total = f.totalItemCount || 0;
    const unread = f.unreadItemCount || 0;
    const isInbox = INBOX_NAMES.has(name.toLowerCase());

    if (isInbox) {
      inboxId = f.id;
      inboxCount = total;
      // Detect the "Other" tab — Microsoft's Focused Inbox splits inbox into Focused/Other.
      // In the new Outlook, "Other" can be a physical child folder of Inbox.
      const otherChild = ((f.childFolders || []) as any[]).find(
        (cf: any) => OTHER_INBOX_NAMES.has((cf.displayName || "").toLowerCase())
      );
      if (otherChild) {
        otherInboxId = otherChild.id;
        otherInboxCount = otherChild.totalItemCount || 0;
      }
    }

    // Build children from the expanded data (no extra API calls)
    const children: FolderSnapshot[] = ((f.childFolders || []) as any[]).map((cf: any) => ({
      id: cf.id,
      name: cf.displayName || "Ukjent",
      path: `${name}/${cf.displayName || "Ukjent"}`,
      totalCount: cf.totalItemCount || 0,
      unreadCount: cf.unreadItemCount || 0,
      messages: [],
      children: [],
    }));

    folders.push({
      id: f.id,
      name,
      path: name,
      totalCount: total,
      unreadCount: unread,
      messages: [],
      children,
    });
  }

  onProgress("Mappestruktur klar!", 100);
  return { folders, inboxId, inboxCount, otherInboxId, otherInboxCount };
}

export function resetInboxPagination(): void {
  // Kept for compatibility with existing callers.
  // Batch fetching now always re-queries oldest messages from the top,
  // which is stable while messages are moved/deleted between rounds.
}

/** Fetch a batch of messages from inbox (oldest first to ensure sequential processing).
 *  When the main inbox (Focused tab) is empty and an "Other" folder is provided,
 *  falls back to fetching from the Other folder so Focused Inbox is fully processed.
 *  Re-queries from the top each round because moved/deleted messages invalidate cursors. */
export async function fetchInboxBatch(
  inboxFolderId: string,
  top: number = 50,
  otherFolderId?: string,
): Promise<MailMessage[]> {
  const msgRes = await graphGet(
    `/mailFolders/${inboxFolderId}/messages?$top=${top}&$orderby=receivedDateTime asc&$select=id,subject,from,receivedDateTime,isRead`
  );
  const messages = ((msgRes.value || []) as any[]).map((m: any, idx: number) => ({
    id: m.id,
    index: idx,
    subject: m.subject || "(uten emne)",
    fromName: m.from?.emailAddress?.name || "",
    fromAddress: m.from?.emailAddress?.address || "",
    date: m.receivedDateTime || "",
    isRead: m.isRead ?? true,
    folderId: inboxFolderId,
    folderName: "Innboks",
  }));

  // If main inbox is empty and an "Other" folder exists, process it next.
  if (messages.length === 0 && otherFolderId) {
    try {
      const otherRes = await graphGet(
        `/mailFolders/${otherFolderId}/messages?$top=${top}&$orderby=receivedDateTime asc&$select=id,subject,from,receivedDateTime,isRead`
      );
      const otherMessages = ((otherRes.value || []) as any[]).map((m: any, idx: number) => ({
        id: m.id,
        index: idx,
        subject: m.subject || "(uten emne)",
        fromName: m.from?.emailAddress?.name || "",
        fromAddress: m.from?.emailAddress?.address || "",
        date: m.receivedDateTime || "",
        isRead: m.isRead ?? true,
        folderId: otherFolderId,
        folderName: "Innboks / Annet",
      }));

      return otherMessages;
    } catch { /* ignore if Other folder is inaccessible */ }
  }

  return messages;
}

/**
 * Bulk-archive emails from a specific folder that are older than a given date.
 * Works directly via Graph API filtering — no scan/AI needed.
 * Moves in batches of 20, reports progress.
 */
export async function archiveOldFromFolder(
  folderDisplayName: string,
  beforeDate: Date,
  onProgress: (msg: string, moved: number, total: number | null) => void
): Promise<{ moved: number; errors: number }> {
  onProgress("Finner mapper...", 0, null);

  // Find the source folder
  const foldersRes = await graphGet("/mailFolders?$top=100");
  const rawFolders: any[] = foldersRes.value || [];
  const sourceFolder = rawFolders.find(
    (f: any) => (f.displayName || "").toLowerCase() === folderDisplayName.toLowerCase()
  );
  if (!sourceFolder) throw new Error(`Fant ikke mappen "${folderDisplayName}"`);

  // Find or create Archive folder
  let archiveFolder = rawFolders.find(
    (f: any) => (f.displayName || "").toLowerCase() === "arkiv" || (f.displayName || "").toLowerCase() === "archive"
  );
  let archiveId: string;
  if (archiveFolder) {
    archiveId = archiveFolder.id;
  } else {
    archiveId = await createFolder("Arkiv");
  }

  const dateFilter = beforeDate.toISOString();
  const sourceName = sourceFolder.displayName;
  const sourceTotal = sourceFolder.totalItemCount || 0;
  onProgress(`Fant "${sourceName}" (${sourceTotal} e-poster). Henter gamle e-poster...`, 0, null);

  let moved = 0;
  let errors = 0;
  const BATCH = 20;

  // Query old emails in batches using server-side filter
  let nextPath: string | null =
    `/mailFolders/${sourceFolder.id}/messages?$filter=receivedDateTime lt ${dateFilter}&$orderby=receivedDateTime asc&$top=${BATCH}&$select=id,subject,receivedDateTime`;

  while (nextPath) {
    const res = await graphGet(nextPath);
    const msgs: any[] = res.value || [];
    if (msgs.length === 0) break;

    for (const m of msgs) {
      try {
        await graphPost(`/messages/${m.id}/move`, { destinationId: archiveId });
        moved++;
      } catch {
        errors++;
      }
      if ((moved + errors) % 5 === 0 || moved + errors === msgs.length) {
        onProgress(
          `Flytter fra "${sourceName}" til Arkiv...`,
          moved,
          null
        );
      }
    }

    // Follow pagination — but since we moved the messages, re-query same filter
    // (moved messages are gone from source, so next page is effectively the same query)
    const raw = res["@odata.nextLink"];
    if (raw && msgs.length === BATCH) {
      // Re-query from the start since moved messages are gone
      nextPath = `/mailFolders/${sourceFolder.id}/messages?$filter=receivedDateTime lt ${dateFilter}&$orderby=receivedDateTime asc&$top=${BATCH}&$select=id,subject,receivedDateTime`;
    } else if (msgs.length === BATCH) {
      // No nextLink but we got a full batch — there might be more
      nextPath = `/mailFolders/${sourceFolder.id}/messages?$filter=receivedDateTime lt ${dateFilter}&$orderby=receivedDateTime asc&$top=${BATCH}&$select=id,subject,receivedDateTime`;
    } else {
      nextPath = null;
    }
  }

  onProgress(`Ferdig! ${moved} e-poster arkivert.`, moved, moved + errors);
  return { moved, errors };
}

/**
 * Find a mail folder by display name, searching recursively through all subfolders.
 * Supports path syntax like "Innboks/Nyhetsbrev" or a plain name like "Nyhetsbrev".
 */
async function findFolderByPath(nameOrPath: string): Promise<string | null> {
  const target = nameOrPath.trim().toLowerCase();
  const parts = target.split("/").filter(Boolean);

  const foldersRes = await graphGet("/mailFolders?$top=100");
  const topFolders: any[] = foldersRes.value || [];

  // If path has multiple parts, resolve step by step: "Innboks/Kunder/Fremtind"
  if (parts.length > 1) {
    let parentFolders = topFolders;
    let currentId: string | null = null;
    for (const part of parts) {
      const match = parentFolders.find(
        (f: any) => (f.displayName || "").toLowerCase() === part
      );
      if (!match) return null;
      currentId = match.id;
      // Fetch children for next level
      const childRes = await graphGet(`/mailFolders/${currentId}/childFolders?$top=100`);
      parentFolders = childRes.value || [];
    }
    return currentId;
  }

  // Single name — first check top-level
  const topMatch = topFolders.find(
    (f: any) => (f.displayName || "").toLowerCase() === target
  );
  if (topMatch) return topMatch.id;

  // Search recursively in subfolders
  for (const f of topFolders) {
    const found = await searchChildFolders(f.id, target);
    if (found) return found;
  }
  return null;
}

async function searchChildFolders(parentId: string, target: string): Promise<string | null> {
  try {
    const res = await graphGet(`/mailFolders/${parentId}/childFolders?$top=100`);
    const children: any[] = res.value || [];
    for (const c of children) {
      if ((c.displayName || "").toLowerCase() === target) return c.id;
    }
    // Go deeper
    for (const c of children) {
      const found = await searchChildFolders(c.id, target);
      if (found) return found;
    }
  } catch { /* permission or no children */ }
  return null;
}

/**
 * Organize emails in a folder into sub-folders grouped by sender.
 * Senders with more messages than `threshold` get their own sub-folder.
 */
export async function organizeByFolder(
  folderDisplayName: string,
  threshold: number,
  onProgress: (msg: string, subFolders: number, moved: number) => void
): Promise<{ subFoldersCreated: number; messagesMoved: number }> {
  onProgress("Finner mappe...", 0, 0);

  // Find the source folder — supports nested paths like "Innboks/Nyhetsbrev"
  const sourceId = await findFolderByPath(folderDisplayName);
  if (!sourceId) throw new Error(`Fant ikke mappen "${folderDisplayName}"`);
  onProgress(`Skanner "${folderDisplayName}"...`, 0, 0);

  // Fetch all messages (up to 1000) to analyze sender frequency
  const allMessages: { id: string; fromAddress: string; fromName: string }[] = [];
  let nextPath: string | null =
    `/mailFolders/${sourceId}/messages?$top=100&$orderby=receivedDateTime desc&$select=id,from`;

  while (nextPath && allMessages.length < 1000) {
    const res = await graphGet(nextPath);
    const msgs: any[] = res.value || [];
    if (msgs.length === 0) break;

    for (const m of msgs) {
      allMessages.push({
        id: m.id,
        fromAddress: (m.from?.emailAddress?.address || "").toLowerCase(),
        fromName: m.from?.emailAddress?.name || "",
      });
    }

    const raw = res["@odata.nextLink"];
    nextPath = raw ? raw.replace("https://graph.microsoft.com/v1.0/me", "") : null;
  }

  onProgress(`Analyserer ${allMessages.length} e-poster...`, 0, 0);

  // Group by sender email
  const senderMap: Record<string, { name: string; ids: string[] }> = {};
  for (const m of allMessages) {
    if (!m.fromAddress) continue;
    if (!senderMap[m.fromAddress]) {
      senderMap[m.fromAddress] = { name: m.fromName || m.fromAddress, ids: [] };
    }
    senderMap[m.fromAddress].ids.push(m.id);
  }

  // Get existing child folders to avoid duplicates
  const childRes = await graphGet(`/mailFolders/${sourceId}/childFolders?$top=100`);
  const existingChildren: Record<string, string> = {};
  for (const cf of (childRes.value || []) as any[]) {
    existingChildren[(cf.displayName || "").toLowerCase()] = cf.id;
  }

  // Create sub-folders and move messages for senders above threshold
  let subFoldersCreated = 0;
  let messagesMoved = 0;

  const senders = Object.entries(senderMap)
    .filter(([_, v]) => v.ids.length > threshold)
    .sort((a, b) => b[1].ids.length - a[1].ids.length);

  for (const [email, sender] of senders) {
    // Clean name for folder display (no slashes or control chars)
    const cleanName = (sender.name || email)
      .replace(/[/\\<>:"|?*]/g, "-")
      .replace(/\s+/g, " ")
      .trim()
      .substring(0, 60);

    // Find existing or create new sub-folder
    let childId = existingChildren[cleanName.toLowerCase()];
    if (!childId) {
      childId = await createFolder(cleanName, sourceId);
      existingChildren[cleanName.toLowerCase()] = childId;
      subFoldersCreated++;
    }

    // Move messages in small batches
    for (const msgId of sender.ids) {
      try {
        await graphPost(`/messages/${msgId}/move`, { destinationId: childId });
        messagesMoved++;
      } catch { /* continue */ }
      if (messagesMoved % 5 === 0) {
        onProgress(
          `Organiserer: ${cleanName} (${sender.ids.length} e-poster)`,
          subFoldersCreated,
          messagesMoved
        );
      }
      if (sender.ids.length > 3) await delay(200); // throttle
    }
  }

  onProgress(`Ferdig!`, subFoldersCreated, messagesMoved);
  return { subFoldersCreated, messagesMoved };
}
