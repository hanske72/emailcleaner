import { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } from "@google/generative-ai";
import type { MailboxSnapshot, FolderSnapshot, MailMessage } from "./outlook";

let genai: GoogleGenerativeAI | null = null;

export function getApiKey(): string | null {
  return localStorage.getItem("gemini-api-key");
}

export function setApiKey(key: string): void {
  localStorage.setItem("gemini-api-key", key);
  genai = null;
}

export function hasApiKey(): boolean {
  return !!getApiKey();
}

const safetySettings = [
  { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
  { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
];

function getModel(thinkingBudget: number = 1024) {
  if (!genai) {
    const key = getApiKey();
    if (!key) throw new Error("Gemini API-nøkkel mangler. Åpne innstillinger.");
    genai = new GoogleGenerativeAI(key);
  }
  const cfg: any = {
    model: "gemini-2.5-flash",
    safetySettings,
    generationConfig: { maxOutputTokens: 16384 },
  };
  // Limit thinking to save output tokens for actual JSON
  try { cfg.generationConfig.thinkingConfig = { thinkingBudget }; } catch (_) {}
  return genai.getGenerativeModel(cfg);
}

/** Try to repair truncated JSON from Gemini */
function repairJson(text: string): string {
  let s = text.trim();
  // Remove markdown fences if present
  s = s.replace(/^```json?\n?/i, "").replace(/\n?```$/i, "");
  // Try parsing as-is first
  try { JSON.parse(s); return s; } catch (_) {}
  // Remove trailing comma before attempting to close
  s = s.replace(/,\s*$/, "");
  // Count open/close brackets
  let openBraces = 0, openBrackets = 0;
  let inString = false, escape = false;
  for (const ch of s) {
    if (escape) { escape = false; continue; }
    if (ch === '\\' && inString) { escape = true; continue; }
    if (ch === '"') { inString = !inString; continue; }
    if (inString) continue;
    if (ch === '{') openBraces++;
    else if (ch === '}') openBraces--;
    else if (ch === '[') openBrackets++;
    else if (ch === ']') openBrackets--;
  }
  // If we're inside a string, close it
  if (inString) s += '"';
  // Close open brackets and braces
  while (openBrackets > 0) { s = s.replace(/,\s*$/, ""); s += ']'; openBrackets--; }
  while (openBraces > 0) { s = s.replace(/,\s*$/, ""); s += '}'; openBraces--; }
  return s;
}

export interface CleanupAction {
  type: "create_folder" | "move" | "archive" | "delete";
  description: string;
  messageIndices: number[];
  targetFolder?: string;
  newFolderName?: string;
  reason: string;
  approved: boolean;
}

export interface CleanupPlan {
  summary: string;
  observations: string[];
  actions: CleanupAction[];
}

export async function analyzeMailbox(
  snapshot: MailboxSnapshot,
  log?: (msg: string) => void,
  onBatchProgress?: (batch: number, totalBatches: number, analyzedSoFar: number, totalMessages: number) => void,
): Promise<CleanupPlan> {
  const model = getModel();
  const _log = log || ((msg: string) => console.log("[Gemini]", msg));

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

  // --- Helper: build a COMPACT tree with 1 sample per user folder ---
  function buildFolderTree(folder: any, indent: number): string {
    const prefix = "  ".repeat(indent);
    const isSystem = SYSTEM_FOLDER_NAMES.has(folder.name.toLowerCase());
    const isInbox = INBOX_NAMES.has(folder.name.toLowerCase());
    const lines: string[] = [];

    if (isInbox && indent === 0) {
      lines.push(`${prefix}📥 ${folder.name} (${folder.totalCount}) — innhold nedenfor`);
    } else if (isSystem) {
      lines.push(`${prefix}${folder.name}: ${folder.totalCount}`);
    } else {
      const sample = folder.messages?.[0];
      const hint = sample ? ` eks: "${(sample.subject || "").substring(0, 40)}"` : "";
      lines.push(`${prefix}📁 ${folder.path || folder.name} (${folder.totalCount})${hint}`);
    }

    if (folder.children && folder.children.length > 0) {
      for (const child of folder.children) {
        lines.push(buildFolderTree(child, indent + 1));
      }
    }
    return lines.join("\n");
  }

  // Build complete folder structure
  const folderTreeLines = snapshot.folders.map((f) => buildFolderTree(f, 0));
  let folderTree = folderTreeLines.join("\n");
  if (folderTree.length > 4000) {
    folderTree = folderTree.substring(0, 4000) + "\n  ... (avkortet)";
  }

  // Find inbox and its messages
  const inboxFolder = snapshot.folders.find((f) => INBOX_NAMES.has(f.name.toLowerCase()) && f.messages.length > 0);
  const allInboxMessages = inboxFolder ? inboxFolder.messages : [];

  // Collect all user folder paths
  function collectUserPaths(folder: any): string[] {
    const paths: string[] = [];
    const isSystem = SYSTEM_FOLDER_NAMES.has(folder.name.toLowerCase());
    const isInbox = INBOX_NAMES.has(folder.name.toLowerCase());
    if (!isSystem && !isInbox && folder.totalCount > 0) {
      paths.push(folder.path || folder.name);
    }
    if (folder.children) {
      for (const child of folder.children) {
        paths.push(...collectUserPaths(child));
      }
    }
    return paths;
  }
  const allUserPaths = snapshot.folders.flatMap((f) => {
    if (INBOX_NAMES.has(f.name.toLowerCase()) && f.children) {
      return f.children.flatMap((c: any) => collectUserPaths(c));
    }
    return collectUserPaths(f);
  });

  // --- Batch processing ---
  const BATCH_SIZE = 50;
  const totalMessages = allInboxMessages.length;
  const totalBatches = Math.max(1, Math.ceil(totalMessages / BATCH_SIZE));

  _log(`Innboks har ${totalMessages} meldinger — delt i ${totalBatches} batch(er) à ${BATCH_SIZE}`);

  const allActions: CleanupAction[] = [];
  const allObservations: string[] = [];
  let combinedSummary = "";

  for (let batchIdx = 0; batchIdx < totalBatches; batchIdx++) {
    const batchStart = batchIdx * BATCH_SIZE;
    const batchMessages = allInboxMessages.slice(batchStart, batchStart + BATCH_SIZE);
    const batchNum = batchIdx + 1;

    _log(`--- Batch ${batchNum}/${totalBatches}: meldinger ${batchStart}–${batchStart + batchMessages.length - 1} ---`);
    if (onBatchProgress) onBatchProgress(batchNum, totalBatches, batchStart, totalMessages);

    // Domain frequency for this batch
    const domainCounts = new Map<string, number>();
    for (const m of batchMessages) {
      const domain = (m.fromAddress || "").split("@")[1]?.toLowerCase();
      if (domain) domainCounts.set(domain, (domainCounts.get(domain) || 0) + 1);
    }
    const topDomains = [...domainCounts.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([domain, count]) => `  ${domain}: ${count}`)
      .join("\n");

    // Build inbox listing for this batch
    let inboxText = "";
    if (batchMessages.length > 0) {
      const header = totalBatches > 1
        ? `Innboks — batch ${batchNum}/${totalBatches} (viser meldinger ${batchStart}–${batchStart + batchMessages.length - 1} av ${totalMessages} totalt)`
        : `Innboks (${inboxFolder!.totalCount} totalt, ${inboxFolder!.unreadCount} uleste, viser ${batchMessages.length})`;
      const msgs = batchMessages
        .map((m, localIdx) => {
          const subj = m.subject && m.subject.length > 60 ? m.subject.substring(0, 60) + "…" : m.subject;
          return `[${localIdx}] "${subj}" fra ${m.fromAddress}`;
        })
        .join("\n");
      inboxText = `${header}\n${msgs}`;
    } else {
      inboxText = "Innboksen er tom.";
    }

    const batchContext = totalBatches > 1
      ? `\nDette er batch ${batchNum} av ${totalBatches}. Analyser KUN meldingene i denne batchen.\n`
      : "";

    const prompt = `Du er en e-postorganiserings-ekspert. Du skal først studere mappestrukturen nøye, deretter analysere innboksen og foreslå hvor e-poster bør flyttes.
${batchContext}
STEG 1 — FORSTÅ MAPPESTRUKTUREN:
Brukeren har følgende mappestruktur med undermapper. Eksempler på innhold vises der det er tilgjengelig.

${folderTree}

GYLDIGE FLYTTEMÅL (eksisterende brukermapper):
${allUserPaths.length > 0 ? allUserPaths.map((p) => `  📁 ${p}`).join("\n") : "  (ingen egendefinerte mapper)"}

STEG 2 — DOMENE-ANALYSE:
${topDomains || "  (ingen data)"}

STEG 3 — ANALYSER INNBOKSEN:
${inboxText}

OPPGAVE:
Lag en KONKRET ryddeplan. ALLE meldingene i listen MÅ være med i minst én handling — ingen melding skal utelates.

REGLER:
- Skriv ALT på norsk
- KRITISK: HVER ENESTE melding i listen MÅ inkluderes i en handling. Telle opp og sjekk at alle [indeks]-numre er dekket.
- Fokuser KUN på innboksen
- BRUK EKSISTERENDE MAPPER som flyttemål der innholdet passer
- Inkluder full mappesti som targetFolder (f.eks. "Innboks/_1_Kunder/Fremtind")
- Bruk domene-analysen til å gruppere e-poster fra samme avsender
- Henvis til e-poster med NØYAKTIG de [indeks]-numrene som står i listen
- Grupper lignende e-poster (nyhetsbrev, reklame, fakturaer)
- Opprett nye mapper for hyppige avsendere som mangler mappe
- Bruk så mange handlinger som trengs — ingen øvre grense
- "archive" = flytt til Arkiv, "delete" = flytt til Slettet
- Opprett mapper (create_folder) FØR move-handlinger som bruker dem
- Aldri slett noe som kan være viktig
- Meldinger som ikke passer noe sted kan arkiveres

SVAR-FORMAT (eksakt JSON):
{
  "summary": "kort oppsummering av innboksen",
  "observations": ["observasjon 1", "observasjon 2"],
  "actions": [
    {
      "type": "move",
      "description": "Flytt e-poster fra X til Y",
      "messageIndices": [0, 1, 2],
      "targetFolder": "Innboks/_1_Kunder/Firmanavn",
      "reason": "Disse hører til Firmanavn"
    }
  ]
}
Gyldige type-verdier: "create_folder", "move", "archive", "delete"
For create_folder bruk "newFolderName" istedenfor "targetFolder".
Svar BARE med JSON — ingen tekst før eller etter.`;

    _log(`Prompt: ${prompt.length} tegn, ${batchMessages.length} meldinger`);

    let batchPlan!: CleanupPlan;
    let lastError = "";
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        _log(`Forsøk ${attempt}/3 — sender til Gemini...`);
        const result = await model.generateContent(prompt);
        const resp = result.response;
        const finishReason = resp.candidates?.[0]?.finishReason || "ukjent";
        const blockReason = resp.promptFeedback?.blockReason;
        _log(`Svar mottatt: finishReason=${finishReason}${blockReason ? `, BLOKKERT: ${blockReason}` : ""}`);

        if (blockReason) {
          throw new Error(`Gemini blokkerte forespørselen: ${blockReason}`);
        }

        let raw: string;
        try {
          raw = resp.text();
        } catch (textErr: any) {
          _log(`resp.text() feilet: ${textErr.message}`);
          lastError = `text() feilet: ${textErr.message}`;
          if (attempt === 3) throw new Error(`Gemini returnerte tomt/ugyldig svar: ${textErr.message}`);
          continue;
        }

        _log(`Rå respons: ${raw.length} tegn. Start: ${raw.substring(0, 150)}...`);

        if (finishReason === "MAX_TOKENS") {
          _log("ADVARSEL: Respons ble avkortet (MAX_TOKENS)");
        }

        const text = repairJson(raw);
        batchPlan = JSON.parse(text) as CleanupPlan;
        // Remap batch-local indices (0-based within batch) to global message indices
        if (batchPlan.actions) {
          for (const action of batchPlan.actions) {
            if (action.messageIndices) {
              action.messageIndices = action.messageIndices
                .filter(localIdx => localIdx >= 0 && localIdx < batchMessages.length)
                .map(localIdx => batchMessages[localIdx].index);
            }
          }
        }
        _log(`Batch ${batchNum} parse OK! ${batchPlan.actions?.length || 0} handlinger funnet.`);
        break;
      } catch (e: any) {
        if (e.message?.includes("blokkerte")) throw e;
        lastError = e.message || String(e);
        _log(`Feil forsøk ${attempt}: ${lastError}`);
        if (attempt === 3) {
          _log(`Batch ${batchNum} feilet etter 3 forsøk — hopper over denne batchen`);
          batchPlan = { summary: "", observations: [`Batch ${batchNum} feilet: ${lastError}`], actions: [] };
        }
      }
    }

    // Accumulate results
    if (batchPlan.summary) {
      combinedSummary = combinedSummary
        ? `${combinedSummary} | ${batchPlan.summary}`
        : batchPlan.summary;
    }
    if (batchPlan.observations) allObservations.push(...batchPlan.observations);
    if (batchPlan.actions) allActions.push(...batchPlan.actions);
  }

  // Report final progress
  if (onBatchProgress) onBatchProgress(totalBatches, totalBatches, totalMessages, totalMessages);

  // Deduplicate create_folder actions (same folder name)
  const seenFolders = new Set<string>();
  const deduped: CleanupAction[] = [];
  for (const a of allActions) {
    if (a.type === "create_folder") {
      const key = (a.newFolderName || "").toLowerCase();
      if (seenFolders.has(key)) continue;
      seenFolders.add(key);
    }
    deduped.push(a);
  }

  // Find inbox messages not covered by any action — add catch-all
  const coveredIndices = new Set<number>();
  for (const a of deduped) {
    if (a.messageIndices) for (const idx of a.messageIndices) coveredIndices.add(idx);
  }
  const uncovered = allInboxMessages.filter(m => !coveredIndices.has(m.index));
  if (uncovered.length > 0) {
    _log(`${uncovered.length} meldinger ikke dekket av Gemini — legger til catch-all arkivering`);
    deduped.push({
      type: "archive",
      description: `Arkiver ${uncovered.length} resterende meldinger`,
      messageIndices: uncovered.map(m => m.index),
      reason: "Disse meldingene ble ikke inkludert i noen sorteringshandling",
      approved: false, // user must opt-in
    });
  }

  const plan: CleanupPlan = {
    summary: totalBatches > 1
      ? `${combinedSummary} (${totalMessages} meldinger analysert i ${totalBatches} batcher)`
      : combinedSummary,
    observations: allObservations,
    actions: deduped.map((a) => ({ ...a, approved: a.approved ?? true })),
  };

  _log(`Totalt: ${plan.actions.length} handlinger fra ${totalBatches} batch(er)`);
  return plan;
}

/**
 * Analyze a single batch of messages and return sort actions.
 * Used by v4.0 auto-clean: fetch → analyze → execute → repeat.
 */
export async function analyzeBatch(
  messages: MailMessage[],
  folders: FolderSnapshot[],
  batchNum: number,
  inboxRemaining: number,
  log?: (msg: string) => void,
): Promise<CleanupAction[]> {
  const _log = log || (() => {});
  const fastMode = messages.length <= 10 || inboxRemaining <= 20;
  const model = getModel(fastMode ? 128 : 512);
  const maxAttempts = fastMode ? 1 : 3;
  const modelTimeoutMs = fastMode ? 20000 : 45000;

  const _INBOX = new Set(["innboks", "inbox"]);
  const _SYSTEM = new Set([
    "innboks", "inbox", "sendte elementer", "sent items",
    "kladd", "drafts", "utboks", "outbox",
    "slettede elementer", "deleted items",
    "s\u00f8ppelpost", "junk email", "junk", "spam",
    "arkiv", "archive", "samtalelogg", "conversation history",
  ]);

  function _tree(f: any, indent: number): string {
    const pre = "  ".repeat(indent);
    const sys = _SYSTEM.has(f.name.toLowerCase());
    const inb = _INBOX.has(f.name.toLowerCase());
    const lines: string[] = [];
    if (inb && indent === 0) {
      lines.push(`${pre}\ud83d\udce5 ${f.name} (${f.totalCount})`);
    } else if (sys) {
      lines.push(`${pre}${f.name}: ${f.totalCount}`);
    } else {
      const sample = f.messages?.[0];
      const hint = sample ? ` eks: "${(sample.subject || "").substring(0, 40)}"` : "";
      lines.push(`${pre}\ud83d\udcc1 ${f.path || f.name} (${f.totalCount})${hint}`);
    }
    if (f.children?.length > 0) for (const c of f.children) lines.push(_tree(c, indent + 1));
    return lines.join("\n");
  }

  let folderTree = folders.map(f => _tree(f, 0)).join("\n");
  const folderTreeLimit = fastMode ? 1500 : 4000;
  if (folderTree.length > folderTreeLimit) folderTree = folderTree.substring(0, folderTreeLimit) + "\n... (avkortet)";

  function _paths(f: any): string[] {
    const r: string[] = [];
    if (!_SYSTEM.has(f.name.toLowerCase()) && !_INBOX.has(f.name.toLowerCase()) && f.totalCount > 0) {
      r.push(f.path || f.name);
    }
    if (f.children) for (const c of f.children) r.push(..._paths(c));
    return r;
  }
  const userPaths = folders.flatMap(f => {
    if (_INBOX.has(f.name.toLowerCase()) && f.children) {
      return f.children.flatMap((c: any) => _paths(c));
    }
    return _paths(f);
  });
  const userPathsForPrompt = fastMode ? userPaths.slice(0, 20) : userPaths;

  const domainCounts = new Map<string, number>();
  for (const m of messages) {
    const d = (m.fromAddress || "").split("@")[1]?.toLowerCase();
    if (d) domainCounts.set(d, (domainCounts.get(d) || 0) + 1);
  }
  const topDomains = [...domainCounts.entries()]
    .sort((a, b) => b[1] - a[1]).slice(0, 10)
    .map(([d, c]) => `  ${d}: ${c}`).join("\n");

  const msgList = messages.map((m, i) => {
    const s = m.subject?.length > 60 ? m.subject.substring(0, 60) + "\u2026" : m.subject;
    return `[${i}] "${s}" fra ${m.fromAddress}`;
  }).join("\n");

  const prompt = `Du er en e-postorganiserings-ekspert. Sorter disse ${messages.length} e-postene i mapper.

MAPPESTRUKTUR:
${folderTree}

GYLDIGE FLYTTEM\u00c5L (eksisterende brukermapper):
${userPathsForPrompt.length > 0 ? userPathsForPrompt.map(p => `  \ud83d\udcc1 ${p}`).join("\n") : "  (ingen egendefinerte mapper)"}

DOMENE-ANALYSE:
${topDomains || "  (ingen data)"}

E-POSTER \u00c5 SORTERE (batch ${batchNum}, ${messages.length} av ${inboxRemaining} gjenv\u00e6rende):
${msgList}

OPPGAVE:
Sorter ALLE ${messages.length} meldingene. HVER melding M\u00c5 inkluderes i en handling.

REGLER:
- Skriv ALT p\u00e5 norsk
- KRITISK: HVER ENESTE melding M\u00c5 inkluderes. Sjekk at alle indekser 0\u2013${messages.length - 1} er dekket.
- BRUK EKSISTERENDE MAPPER der innholdet passer
- Inkluder full mappesti som targetFolder (f.eks. "Innboks/_1_Kunder/Fremtind")
- Grupper e-poster fra samme avsender/domene
- Opprett nye mapper for hyppige avsendere som mangler mappe
- "archive" = flytt til Arkiv, "delete" = flytt til Slettet
- Opprett mapper (create_folder) F\u00d8R move-handlinger som bruker dem
- Aldri slett noe som kan v\u00e6re viktig
- Meldinger som ikke passer noe sted: arkiver dem

SVAR-FORMAT (eksakt JSON):
{
  "actions": [
    { "type": "create_folder", "description": "Opprett mappe X", "messageIndices": [], "newFolderName": "Innboks/NyMappe", "reason": "..." },
    { "type": "move", "description": "Flytt til X", "messageIndices": [0, 1, 2], "targetFolder": "Innboks/_1_Kunder/Firma", "reason": "..." },
    { "type": "archive", "description": "Arkiver gamle", "messageIndices": [3, 4], "reason": "..." }
  ]
}
Gyldige type-verdier: "create_folder", "move", "archive", "delete"
Svar BARE med JSON \u2014 ingen tekst f\u00f8r eller etter.`;

  _log(`Batch ${batchNum}: ${messages.length} meldinger, prompt ${prompt.length} tegn${fastMode ? " (fast mode)" : ""}`);

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      _log(`Batch ${batchNum} fors\u00f8k ${attempt}/${maxAttempts}...`);
      const result = await Promise.race([
        model.generateContent(prompt),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error(`Gemini timeout etter ${Math.round(modelTimeoutMs / 1000)}s`)), modelTimeoutMs)
        ),
      ]);
      const resp = result.response;
      const finishReason = resp.candidates?.[0]?.finishReason || "ukjent";
      const blockReason = resp.promptFeedback?.blockReason;
      if (blockReason) throw new Error(`Gemini blokkerte: ${blockReason}`);

      let raw: string;
      try { raw = resp.text(); } catch (e: any) {
        if (attempt === 3) throw new Error(`Tomt svar: ${e.message}`);
        continue;
      }

      _log(`Batch ${batchNum}: ${raw.length} tegn, finishReason=${finishReason}`);

      const text = repairJson(raw);
      const parsed = JSON.parse(text);
      const actions: CleanupAction[] = (parsed.actions || []).map((a: any) => ({
        ...a,
        messageIndices: (a.messageIndices || []).filter((i: number) => i >= 0 && i < messages.length),
        approved: true,
      }));

      _log(`Batch ${batchNum}: ${actions.length} handlinger`);
      return actions;
    } catch (e: any) {
      if (e.message?.includes("blokkerte")) throw e;
      _log(`Batch ${batchNum} feil ${attempt}: ${e.message}`);
      if (attempt === maxAttempts) {
        _log(`Batch ${batchNum} feilet \u2014 tom liste`);
        return [];
      }
    }
  }
  return [];
}

// --- Chat ---

export interface ChatMessage {
  role: "user" | "model";
  text: string;
}

let chatHistory: ChatMessage[] = [];
let mailboxContext: string = "";

export function initChat(snapshot: MailboxSnapshot | null): void {
  chatHistory = [];
  if (snapshot) {
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

    // Build folder tree for chat context
    function chatFolderTree(folder: any, indent: number): string {
      const prefix = "  ".repeat(indent);
      const isSystem = SYSTEM_FOLDER_NAMES.has(folder.name.toLowerCase());
      const lines: string[] = [];
      if (isSystem) {
        lines.push(`${prefix}${folder.name}: ${folder.totalCount}`);
      } else {
        lines.push(`${prefix}📁 ${folder.path || folder.name}: ${folder.totalCount}`);
        if (folder.messages && folder.messages.length > 0) {
          const samples = folder.messages.slice(0, 3);
          for (const m of samples) {
            const subj = m.subject && m.subject.length > 50 ? m.subject.substring(0, 50) + "…" : m.subject;
            lines.push(`${prefix}  "${subj}" fra ${m.fromName || m.fromAddress}`);
          }
        }
      }
      if (folder.children) {
        for (const child of folder.children) {
          lines.push(chatFolderTree(child, indent + 1));
        }
      }
      return lines.join("\n");
    }
    const folderTree = snapshot.folders.map((f) => chatFolderTree(f, 0)).join("\n");

    const inboxFolder = snapshot.folders.find((f) => INBOX_NAMES.has(f.name.toLowerCase()) && f.messages.length > 0);
    let inboxText = "Innboksen er tom.";
    if (inboxFolder && inboxFolder.messages.length > 0) {
      const header = `📥 ${inboxFolder.name} (${inboxFolder.totalCount} totalt, ${inboxFolder.unreadCount} uleste)`;
      const msgs = inboxFolder.messages.slice(0, 100)
        .map((m) => {
          const subj = m.subject && m.subject.length > 60 ? m.subject.substring(0, 60) + "…" : m.subject;
          return `  [${m.index}] "${subj}" fra ${m.fromName || m.fromAddress}`;
        })
        .join("\n");
      inboxText = `${header}\n${msgs}`;
    }

    mailboxContext = `POSTBOKS-KONTEKST:\n${snapshot.totalMessages} e-poster totalt, ${snapshot.totalUnread} uleste.\n\nMAPPESTRUKTUR:\n${folderTree}\n\nINNBOKS:\n${inboxText}`;
  } else {
    mailboxContext = "";
  }
}

export function getChatHistory(): ChatMessage[] {
  return [...chatHistory];
}

export async function sendChatMessage(userMessage: string): Promise<string> {
  const model = getModel();

  chatHistory.push({ role: "user", text: userMessage });

  const systemPrompt = `Du er en hjelpsom e-postassistent i appen "E-postrydder". Svar alltid på norsk. Vær kort og konkret.
${mailboxContext ? "\nBrukeren har skannet postboksen sin. Her er dataene:\n" + mailboxContext : "\nBrukeren har ikke skannet postboksen ennå."}

Du kan hjelpe med:
- Spørsmål om e-postene i postboksen
- Tips for e-postorganisering
- Forklare hva appen gjør
- Generelle spørsmål

Svar i ren tekst (ikke JSON). Bruk gjerne emoji for lesbarhet.`;

  const contents = [
    { role: "user" as const, parts: [{ text: systemPrompt }] },
    { role: "model" as const, parts: [{ text: "Forstått! Jeg er klar til å hjelpe deg med e-posten din. Hva lurer du på?" }] },
    ...chatHistory.map((m) => ({
      role: m.role === "user" ? ("user" as const) : ("model" as const),
      parts: [{ text: m.text }],
    })),
  ];

  const result = await model.generateContent({ contents });
  const reply = result.response.text().trim();

  chatHistory.push({ role: "model", text: reply });
  return reply;
}
