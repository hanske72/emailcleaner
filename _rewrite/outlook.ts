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
  totalCount: number;
  unreadCount: number;
  messages: MailMessage[];
}

export interface MailboxSnapshot {
  folders: FolderSnapshot[];
  totalMessages: number;
  totalUnread: number;
  allMessages: MailMessage[];
}

// Helper: read property regardless of PascalCase or camelCase
function p(obj: any, name: string): any {
  if (!obj) return undefined;
  return obj[name] ?? obj[name.charAt(0).toLowerCase() + name.slice(1)];
}

async function getAuth(): Promise<{ token: string; baseUrl: string }> {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true } as any, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const restUrl = Office.context.mailbox.restUrl;
        // Handle both Outlook REST API and Microsoft Graph endpoints
        const baseUrl = restUrl.includes("graph.microsoft.com")
          ? `${restUrl}/v1.0/me`
          : `${restUrl}/v2.0/me`;
        resolve({ token: result.value, baseUrl });
      } else {
        reject(new Error(result.error?.message || "Kunne ikke koble til postboks. Prøv å lukke og åpne tillegget."));
      }
    });
  });
}

async function restGet(baseUrl: string, token: string, path: string): Promise<any> {
  const res = await fetch(`${baseUrl}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (res.status === 401) {
    throw new Error("Tilgang nektet. Lukk tillegget og åpne det på nytt.");
  }
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`API-feil ${res.status}: ${text || res.statusText}`);
  }
  return res.json();
}

export async function scanMailbox(
  onProgress: (msg: string, pct: number) => void
): Promise<MailboxSnapshot> {
  onProgress("Kobler til postboks...", 0);
  const { token, baseUrl } = await getAuth();

  onProgress("Henter mapper...", 5);
  const foldersRes = await restGet(baseUrl, token, "/mailfolders?$top=100");
  const rawFolders: any[] = foldersRes.value || [];

  const snapshot: MailboxSnapshot = {
    folders: [],
    totalMessages: 0,
    totalUnread: 0,
    allMessages: [],
  };

  let globalIndex = 0;
  const active = rawFolders.filter((f) => (p(f, "TotalItemCount") || 0) > 0);

  for (let i = 0; i < active.length; i++) {
    const f = active[i];
    const name: string = p(f, "DisplayName") || "Ukjent";
    const total: number = p(f, "TotalItemCount") || 0;
    const unread: number = p(f, "UnreadItemCount") || 0;
    const fid: string = p(f, "Id");
    const pct = 5 + Math.round(((i + 1) / active.length) * 90);
    onProgress(`Skanner ${name} (${total} e-poster)...`, pct);

    let messages: MailMessage[] = [];
    try {
      const msgRes = await restGet(
        baseUrl,
        token,
        `/mailfolders/${fid}/messages?$top=50&$orderby=ReceivedDateTime desc`
      );
      messages = ((msgRes.value || []) as any[]).map((m: any) => {
        const from = p(m, "From");
        const emailAddr = from ? p(from, "EmailAddress") : null;
        return {
          id: p(m, "Id"),
          index: globalIndex++,
          subject: p(m, "Subject") || "(uten emne)",
          fromName: emailAddr ? (p(emailAddr, "Name") || "") : "",
          fromAddress: emailAddr ? (p(emailAddr, "Address") || "") : "",
          date: p(m, "ReceivedDateTime") || "",
          isRead: p(m, "IsRead") ?? true,
          folderId: fid,
          folderName: name,
        } as MailMessage;
      });
    } catch {
      // Skip folders we can't read (e.g., search folders)
    }

    snapshot.folders.push({ id: fid, name, totalCount: total, unreadCount: unread, messages });
    snapshot.totalMessages += total;
    snapshot.totalUnread += unread;
    snapshot.allMessages.push(...messages);
  }

  onProgress("Skanning fullført!", 100);
  return snapshot;
}

export async function moveMessages(ids: string[], targetFolderId: string): Promise<number> {
  const { token, baseUrl } = await getAuth();
  let ok = 0;
  for (const id of ids) {
    try {
      const res = await fetch(`${baseUrl}/messages/${id}/move`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ DestinationId: targetFolderId }),
      });
      if (res.ok) ok++;
    } catch {
      /* continue */
    }
  }
  return ok;
}

export async function createFolder(name: string): Promise<string> {
  const { token, baseUrl } = await getAuth();
  const res = await fetch(`${baseUrl}/mailfolders`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ DisplayName: name }),
  });
  if (!res.ok) throw new Error(`Kunne ikke opprette mappe "${name}"`);
  const data = await res.json();
  return p(data, "Id");
}

export async function deleteMessages(ids: string[]): Promise<number> {
  const { token, baseUrl } = await getAuth();
  let ok = 0;
  for (const id of ids) {
    try {
      const res = await fetch(`${baseUrl}/messages/${id}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${token}` },
      });
      if (res.ok) ok++;
    } catch {
      /* continue */
    }
  }
  return ok;
}
