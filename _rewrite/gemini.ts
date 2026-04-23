import { GoogleGenerativeAI } from "@google/generative-ai";
import type { MailboxSnapshot } from "./outlook";

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

function getModel() {
  if (!genai) {
    const key = getApiKey();
    if (!key) throw new Error("Gemini API-nøkkel mangler. Åpne innstillinger.");
    genai = new GoogleGenerativeAI(key);
  }
  return genai.getGenerativeModel({ model: "gemini-2.0-flash" });
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

export async function analyzeMailbox(snapshot: MailboxSnapshot): Promise<CleanupPlan> {
  const model = getModel();

  // Build folder + message listing for the prompt
  const folderTexts = snapshot.folders.map((f) => {
    const header = `\u{1F4C1} ${f.name} (${f.totalCount} totalt, ${f.unreadCount} uleste)`;
    if (f.messages.length === 0) return header + "\n  (tom eller utilgjengelig)";
    const msgs = f.messages
      .map((m) => {
        const status = m.isRead ? "LEST" : "ULEST";
        const d = m.date
          ? new Date(m.date).toLocaleDateString("nb-NO", { day: "numeric", month: "short", year: "numeric" })
          : "ukjent";
        return `  [${m.index}] "${m.subject}" fra ${m.fromName || m.fromAddress} (${m.fromAddress}) \u2014 ${d} [${status}]`;
      })
      .join("\n");
    return `${header}\n${msgs}`;
  });

  const prompt = `Du er en e-postorganiserings-ekspert. Analyser denne postboksen og lag en konkret ryddeplan.

POSTBOKS:
${snapshot.totalMessages} e-poster totalt, ${snapshot.totalUnread} uleste, ${snapshot.folders.length} mapper.

${folderTexts.join("\n\n")}

OPPGAVE:
1. Beskriv kort hva du observerer (problemer, m\u00f8nstre, rot)
2. Lag en konkret ryddeplan med handlinger

REGLER:
- Skriv ALT p\u00e5 norsk
- V\u00e6r konservativ \u2014 aldri slett noe som kan v\u00e6re viktig
- Henvis til e-poster med [indeks]-nummer fra listen over
- Grupper lignende e-poster (nyhetsbrev, reklame, fakturaer, etc.)
- Foresl\u00e5 nye mapper kun hvis det gir bedre organisering
- Maks 10 handlinger
- "archive" betyr flytt til Arkiv-mappen
- "delete" betyr flytt til Slettet-mappen
- Opprett mapper f\u00f8r du refererer til dem i move-handlinger

Returner KUN gyldig JSON (ingen markdown, ingen komentarer):
{
  "summary": "<2-3 setninger om postboksen og hovedutfordringene>",
  "observations": ["<observasjon 1>", "<observasjon 2>", "..."],
  "actions": [
    {
      "type": "create_folder",
      "description": "<kort beskrivelse>",
      "newFolderName": "<mappenavn>",
      "messageIndices": [],
      "reason": "<hvorfor>"
    },
    {
      "type": "move",
      "description": "<kort beskrivelse>",
      "messageIndices": [1, 5, 12],
      "targetFolder": "<mappenavn>",
      "reason": "<hvorfor>"
    }
  ]
}`;

  const result = await model.generateContent(prompt);
  const text = result.response
    .text()
    .trim()
    .replace(/^```json?\n?/i, "")
    .replace(/\n?```$/i, "");

  const plan = JSON.parse(text) as CleanupPlan;
  // Default all actions to approved
  plan.actions = (plan.actions || []).map((a) => ({ ...a, approved: true }));
  return plan;
}
