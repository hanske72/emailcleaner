import React, { useState, useCallback, useEffect, useRef } from "react";
import {
  hasApiKey,
  setApiKey as saveApiKey,
  getApiKey,
  analyzeBatch,
  CleanupAction,
  initChat,
  sendChatMessage,
  getChatHistory,
  ChatMessage,
} from "../services/gemini";
import {
  scanFolderStructure,
  fetchInboxBatch,
  moveMessages,
  createFolder,
  deleteMessages,
  archiveOldFromFolder,
  organizeByFolder,
  setStrictTokenMode,
  FolderSnapshot,
} from "../services/outlook";
import {
  signIn,
  signOut,
  isSignedIn,
  getSignedInUser,
  tryRestoreSession,
  setDebugLogger,
  ensureTokenValid,
} from "../services/graphAuth";
import "./taskpane.css";

const APP_VERSION = "4.1";

type Phase = "idle" | "working" | "done" | "error" | "quick-archive" | "organize";

function App() {
  const [phase, setPhase] = useState<Phase>("idle");
  const [progress, setProgress] = useState({ message: "", percent: 0 });
  const [error, setError] = useState("");
  const [showSettings, setShowSettings] = useState(!hasApiKey());
  const [apiKey, setApiKeyInput] = useState(getApiKey() || "");
  const [results, setResults] = useState<string[]>([]);
  const [signedInAs, setSignedInAs] = useState("");
  const [signingIn, setSigningIn] = useState(false);
  const [debugLog, setDebugLog] = useState<string[]>([]);
  const [showDebug, setShowDebug] = useState(false);
  const debugRef = useRef<HTMLDivElement>(null);
  const [showChat, setShowChat] = useState(false);
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>([]);
  const [chatInput, setChatInput] = useState("");
  const [chatSending, setChatSending] = useState(false);
  const chatRef = useRef<HTMLDivElement>(null);
  const [quickFolder, setQuickFolder] = useState("Sendte elementer");
  const [quickYears, setQuickYears] = useState(2);
  const [quickMoved, setQuickMoved] = useState(0);
  const [quickMsg, setQuickMsg] = useState("");
  const [orgFolder, setOrgFolder] = useState("Innboks");
  const [orgThreshold, setOrgThreshold] = useState(5);
  const [orgMsg, setOrgMsg] = useState("");
  const [orgSubs, setOrgSubs] = useState(0);
  const [orgMoved, setOrgMoved] = useState(0);

  const addLog = useCallback((msg: string) => {
    const ts = new Date().toLocaleTimeString("nb-NO");
    setDebugLog((prev) => [...prev, `[${ts}] ${msg}`]);
  }, []);

  // Scroll debug log to bottom
  useEffect(() => {
    if (debugRef.current) debugRef.current.scrollTop = debugRef.current.scrollHeight;
  }, [debugLog]);

  // Scroll chat to bottom
  useEffect(() => {
    if (chatRef.current) chatRef.current.scrollTop = chatRef.current.scrollHeight;
  }, [chatMessages]);

  const openChat = () => {
    initChat(null);
    setChatMessages(getChatHistory());
    setShowChat(true);
  };

  const handleSendChat = async () => {
    const msg = chatInput.trim();
    if (!msg || chatSending) return;
    setChatInput("");
    setChatSending(true);
    try {
      const reply = await sendChatMessage(msg);
      setChatMessages(getChatHistory());
    } catch (e: any) {
      setChatMessages((prev) => [...prev, { role: "model", text: `Feil: ${e.message}` }]);
    } finally {
      setChatSending(false);
    }
  };

  // Register logger and restore session on mount
  useEffect(() => {
    setDebugLogger(addLog);
    addLog(`App startet v${APP_VERSION}`);
    addLog(`Office.context: ${typeof Office !== "undefined" && Office.context ? "OK" : "MANGLER"}`);
    tryRestoreSession().then((ok) => {
      if (ok) {
        const user = getSignedInUser() || "(bruker uten e-post)";
        setSignedInAs(user);
        addLog(`Sesjon gjenopprettet: ${user}`);
      } else {
        addLog("Ingen gyldig sesjon funnet");
      }
    }).catch((e) => {
      addLog(`Sesjonsfeil: ${e.message}`);
    });
  }, [addLog]);

  const handleSaveSettings = () => {
    if (apiKey.trim()) saveApiKey(apiKey.trim());
    if (hasApiKey()) setShowSettings(false);
  };

  const handleSignIn = async () => {
    addLog("Bruker klikket 'Logg inn med Microsoft'");
    setSigningIn(true);
    try {
      addLog("Kaller signIn()...");
      const email = await signIn();
      const displayName = email || "(bruker uten e-post)";
      addLog(`Innlogging OK: ${displayName}`);
      setSignedInAs(displayName);
    } catch (e: any) {
      addLog(`Innlogging FEILET: ${e.message}`);
      setError(e.message || "Innlogging feilet");
      setPhase("error");
    } finally {
      setSigningIn(false);
    }
  };

  const startAutoClean = useCallback(async () => {
    if (!hasApiKey()) {
      setShowSettings(true);
      return;
    }
    setPhase("working");
    setResults([]);
    setError("");

    if (!ensureTokenValid()) {
      addLog("Token utløpt");
      setError("Sesjonen din har utløpt. Logg inn på nytt først.");
      setPhase("error");
      return;
    }

    setStrictTokenMode(true);

    try {
      // Step 1: Scan folder structure (once)
      addLog("Skanner mappestruktur...");
      setProgress({ message: "Skanner mappestruktur...", percent: 5 });
      const { folders, inboxId, inboxCount, otherInboxId, otherInboxCount } = await scanFolderStructure((msg, pct) => {
        setProgress({ message: msg, percent: Math.round(pct * 0.1) });
      });

      if (!inboxId) throw new Error("Fant ikke innboksen");
      if (inboxCount === 0 && otherInboxCount === 0) {
        setResults(["\u2705 Innboksen er allerede tom!"]);
        setPhase("done");
        return;
      }
      if (otherInboxId) {
        addLog(`Ser Fokusert: ${inboxCount} + Annet: ${otherInboxCount} e-poster`);
      }

      addLog(`Innboks: ${inboxCount} e-poster`);

      // Build folder ID map (maintained across batches)
      const folderMap: Record<string, string> = {};
      const mapFolders = (flds: any[]) => {
        for (const f of flds) {
          folderMap[f.name.toLowerCase()] = f.id;
          if (f.path) folderMap[f.path.toLowerCase()] = f.id;
          if (f.children?.length > 0) mapFolders(f.children);
        }
      };
      mapFolders(folders);

      const resolveFolderPath = async (path: string): Promise<string> => {
        const cached = folderMap[path.toLowerCase()];
        if (cached) return cached;
        const parts = path.split("/");
        let parentId: string | undefined;
        let builtPath = "";
        for (const part of parts) {
          builtPath = builtPath ? `${builtPath}/${part}` : part;
          let id = folderMap[builtPath.toLowerCase()];
          if (!id) {
            id = await createFolder(part, parentId);
            folderMap[builtPath.toLowerCase()] = id;
            folderMap[part.toLowerCase()] = id;
          }
          parentId = id;
        }
        return parentId!;
      };

      // Step 2: Loop — fetch, analyze, execute, repeat
      let totalProcessed = 0;
      let totalMoved = 0;
      let batchNum = 0;
      const startCount = inboxCount + otherInboxCount;
      let stuckCount = 0;       // detect when no messages are actually moved
      let lastSeenFirstId = "";  // detect same batch returned twice

      while (true) {
        batchNum++;

        setProgress({
          message: `Batch ${batchNum}: Henter e-poster... (${totalProcessed}/${startCount})`,
          percent: Math.min(95, Math.round((totalProcessed / startCount) * 100)),
        });

        const messages = await fetchInboxBatch(inboxId, 50, otherInboxId);
        if (messages.length === 0) {
          addLog("Innboksen er tom!");
          break;
        }

        // Detect stuck loop: same first message means nothing was moved last round
        const firstId = messages[0]?.id || "";
        if (firstId && firstId === lastSeenFirstId) {
          stuckCount++;
          addLog(`Advarsel: Samme e-poster returnert (${stuckCount}/3)`);
          if (stuckCount >= 3) {
            addLog(`Stopper — ${messages.length} e-poster kunne ikke flyttes.`);
            setResults(prev => [...prev, `\u26A0\uFE0F ${messages.length} e-poster i innboksen kunne ikke sorteres.`]);
            break;
          }
        } else {
          stuckCount = 0;
        }
        lastSeenFirstId = firstId;

        addLog(`Batch ${batchNum}: ${messages.length} e-poster`);

        // Analyze with Gemini
        setProgress({
          message: `Batch ${batchNum}: AI analyserer ${messages.length} e-poster...`,
          percent: Math.min(95, Math.round((totalProcessed / startCount) * 100)),
        });

        const actions = await analyzeBatch(messages, folders, batchNum, startCount - totalProcessed, addLog);

        // Execute actions
        let batchMoved = 0;
        setProgress({
          message: `Batch ${batchNum}: Sorterer e-poster...`,
          percent: Math.min(95, Math.round((totalProcessed / startCount) * 100)),
        });

        for (const action of actions) {
          try {
            if (action.type === "create_folder" && action.newFolderName) {
              await resolveFolderPath(action.newFolderName);
            } else if (action.type === "move" && action.targetFolder) {
              const targetId = await resolveFolderPath(action.targetFolder);
              const ids = (action.messageIndices || [])
                .filter((i: number) => i >= 0 && i < messages.length)
                .map((i: number) => messages[i].id);
              if (ids.length > 0) {
                batchMoved += await moveMessages(ids, targetId);
              }
            } else if (action.type === "archive") {
              let archiveId = folderMap["arkiv"] || folderMap["archive"];
              if (!archiveId) {
                archiveId = await createFolder("Arkiv");
                folderMap["arkiv"] = archiveId;
              }
              const ids = (action.messageIndices || [])
                .filter((i: number) => i >= 0 && i < messages.length)
                .map((i: number) => messages[i].id);
              if (ids.length > 0) {
                batchMoved += await moveMessages(ids, archiveId);
              }
            } else if (action.type === "delete") {
              const ids = (action.messageIndices || [])
                .filter((i: number) => i >= 0 && i < messages.length)
                .map((i: number) => messages[i].id);
              if (ids.length > 0) {
                batchMoved += await deleteMessages(ids);
              }
            }
          } catch (e: any) {
            if (e.message === "TOKEN_EXPIRED") throw e;
            const actionName = action.description || action.type || "Ukjent handling";
            addLog(`  \u274C ${actionName}: ${e.message}`);
          }
          if (messages.length > 10) {
            await new Promise(r => setTimeout(r, 25));
          }
        }

        // Catch-all: archive unhandled messages
        const handled = new Set<number>();
        for (const a of actions) {
          if (a.messageIndices) for (const i of a.messageIndices) handled.add(i);
        }
        const unhandledIds = messages
          .filter((_, i) => !handled.has(i))
          .map(m => m.id);
        if (unhandledIds.length > 0) {
          try {
            let archiveId = folderMap["arkiv"] || folderMap["archive"];
            if (!archiveId) {
              archiveId = await createFolder("Arkiv");
              folderMap["arkiv"] = archiveId;
            }
            batchMoved += await moveMessages(unhandledIds, archiveId);
            addLog(`  \u{1F4E6} ${unhandledIds.length} arkivert (catch-all)`);
          } catch (e: any) {
            if (e.message === "TOKEN_EXPIRED") throw e;
            addLog(`  \u26A0\uFE0F Catch-all feilet (${unhandledIds.length} meldinger): ${e.message}`);
            setResults(prev => [...prev, `\u26A0\uFE0F ${unhandledIds.length} e-poster ble ikke flyttet i denne runden (prøver igjen).`]);
          }
        }

        totalMoved += batchMoved;
        totalProcessed += messages.length;
        setResults(prev => [...prev, `\u2705 Batch ${batchNum}: ${batchMoved}/${messages.length} e-poster sortert`]);
        addLog(`Batch ${batchNum}: ${batchMoved} sortert. Totalt: ${totalMoved}/${totalProcessed}`);

        if (messages.length > 20) {
          await new Promise(r => setTimeout(r, 100));
        }
      }

      setProgress({ message: "Ferdig!", percent: 100 });
      setResults(prev => [...prev, `\n\u{1F389} Ferdig! ${totalMoved} e-poster sortert i ${batchNum} batch(er).`]);
      setPhase("done");
    } catch (e: any) {
      addLog(`Feil: ${e.message}`);
      setError(e.message === "TOKEN_EXPIRED"
        ? "Sesjonen utløp. Logg inn på nytt og prøv igjen."
        : (e.message || "Ukjent feil"));
      setPhase("error");
    } finally {
      setStrictTokenMode(false);
    }
  }, [addLog]);

  const reset = () => {
    setPhase("idle");
    setError("");
    setResults([]);
    setProgress({ message: "", percent: 0 });
    setQuickMoved(0);
    setQuickMsg("");
    setOrgSubs(0);
    setOrgMoved(0);
    setOrgMsg("");
  };

  const startQuickArchive = useCallback(async () => {
    setPhase("quick-archive");
    setQuickMoved(0);
    setQuickMsg("Starter...");
    try {
      const cutoff = new Date();
      cutoff.setFullYear(cutoff.getFullYear() - quickYears);
      const result = await archiveOldFromFolder(
        quickFolder,
        cutoff,
        (msg, moved) => {
          setQuickMsg(msg);
          setQuickMoved(moved);
        }
      );
      setQuickMsg(`Ferdig! ${result.moved} e-poster arkivert${result.errors > 0 ? `, ${result.errors} feil` : ""}.`);
    } catch (e: any) {
      setError(e.message || "Ukjent feil");
      setPhase("error");
    }
  }, [quickFolder, quickYears]);

  const startOrganize = useCallback(async () => {
    setPhase("organize");
    setOrgSubs(0);
    setOrgMoved(0);
    setOrgMsg("Starter...");
    try {
      const result = await organizeByFolder(
        orgFolder,
        orgThreshold,
        (msg, subs, moved) => {
          setOrgMsg(msg);
          setOrgSubs(subs);
          setOrgMoved(moved);
        }
      );
      setOrgMsg(`Ferdig! ${result.subFoldersCreated} undermapper opprettet, ${result.messagesMoved} e-poster flyttet.`);
    } catch (e: any) {
      setError(e.message || "Ukjent feil");
      setPhase("error");
    }
  }, [orgFolder, orgThreshold]);

  const approvedCount = 0;

  return (
    <div className="app">
      <header>
        <h1>{"\u{1F4E7}"} E-postrydder</h1>
        <div style={{ display: "flex", gap: "4px", alignItems: "center" }}>
          <span style={{ fontSize: "10px", color: "#999" }}>v{APP_VERSION}</span>
          <button
            className="icon-btn"
            onClick={() => setShowDebug(!showDebug)}
            title="Vis logg"
          >
            {"\u{1F4CB}"}
          </button>
          <button
            className="icon-btn"
            onClick={openChat}
            title="Chat med AI"
          >
            {"\u{1F4AC}"}
          </button>
          <button
            className="icon-btn"
            onClick={() => setShowSettings(!showSettings)}
            title="Innstillinger"
          >
            {"\u2699\uFE0F"}
          </button>
        </div>
      </header>

      {showDebug && (
        <div className="debug-panel">
          <div className="debug-header">
            <strong>Feilsøkingslogg</strong>
            <button onClick={() => setDebugLog([])}>Tøm</button>
          </div>
          <div className="debug-log" ref={debugRef}>
            {debugLog.length === 0 ? (
              <p style={{ color: "#999" }}>Ingen loggmeldinger ennå</p>
            ) : (
              debugLog.map((line, i) => <p key={i}>{line}</p>)
            )}
          </div>
        </div>
      )}

      {showSettings && (
        <div className="settings-panel">
          <label>Gemini API-nøkkel</label>
          <input
            type="password"
            value={apiKey}
            onChange={(e) => setApiKeyInput(e.target.value)}
            placeholder="AIza..."
          />
          <p className="hint">
            Hent nøkkel fra{" "}
            <a href="https://aistudio.google.com/apikey" target="_blank" rel="noopener noreferrer">
              Google AI Studio
            </a>
          </p>
          <button onClick={handleSaveSettings} disabled={!apiKey.trim()}>
            Lagre
          </button>
          {signedInAs && (
            <div style={{ marginTop: "12px", paddingTop: "12px", borderTop: "1px solid #e5e7eb" }}>
              <p style={{ fontSize: "13px", color: "#666", marginBottom: "6px" }}>
                Logget inn som <strong>{signedInAs}</strong>
              </p>
              <button
                className="secondary-btn"
                style={{ width: "100%" }}
                onClick={() => {
                  signOut();
                  setSignedInAs("");
                  setShowSettings(false);
                  reset();
                }}
              >
                {"\u{1F6AA}"} Logg ut
              </button>
            </div>
          )}
        </div>
      )}

      {showChat && (
        <div className="chat-panel">
          <div className="chat-header">
            <strong>{"\u{1F4AC}"} Chat med AI</strong>
            <button onClick={() => setShowChat(false)} title="Lukk">&times;</button>
          </div>
          <div className="chat-messages" ref={chatRef}>
            {chatMessages.length === 0 && !chatSending && (
              <p style={{ color: "#999", fontSize: "13px", textAlign: "center", marginTop: "20px" }}>
                Skriv en melding for å starte samtalen.
              </p>
            )}
            {chatMessages.map((m, i) => (
              <div key={i} className={`chat-msg ${m.role}`}>
                <div className="chat-bubble">{m.text}</div>
              </div>
            ))}
          </div>
          {chatSending && <p className="chat-typing">AI-en tenker...</p>}
          <div className="chat-input-row">
            <input
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSendChat()}
              placeholder="Skriv en melding..."
              disabled={chatSending || !hasApiKey()}
            />
            <button onClick={handleSendChat} disabled={chatSending || !chatInput.trim()}>
              {"\u27A4"}
            </button>
          </div>
        </div>
      )}

      {phase === "idle" && (
        <div className="phase-idle">
          <div className="intro-card">
            <p>
              E-postrydder sorterer innboksen din <strong>automatisk</strong>.
              AI analyserer 50 e-poster om gangen og flytter dem til riktige mapper.
            </p>
            <ul>
              <li>{"\u{1F4C1}"} Skanner mappestruktur</li>
              <li>{"\u{1F9E0}"} AI analyserer i batcher \u00e0 50</li>
              <li>{"\u{1F4E6}"} Flytter automatisk etter analyse</li>
              <li>{"\u{1F504}"} Fortsetter til innboksen er tom</li>
            </ul>
          </div>

          {!hasApiKey() ? (
            <p className="warning">Legg inn Gemini API-nøkkel først {"\u2B06\uFE0F"}</p>
          ) : !signedInAs ? (
            <button className="primary-btn" onClick={handleSignIn} disabled={signingIn}>
              {signingIn ? "Logger inn..." : "\u{1F510} Logg inn med Microsoft"}
            </button>
          ) : (
            <>
              <p style={{ fontSize: "13px", color: "#666", marginBottom: "8px" }}>
                {"\u2705"} Logget inn som <strong>{signedInAs}</strong>
              </p>
              <button className="primary-btn" onClick={startAutoClean}>
                {"\u{1F9F9}"} Start automatisk sortering
              </button>

              <div className="quick-section">
                <h3>{"\u26A1"} Hurtigrydding</h3>
                <p className="hint">Flytt gamle e-poster direkte til Arkiv — uten AI-analyse. Fungerer med alle e-poster, ikke bare de 500 første.</p>
                <div className="quick-form">
                  <label>Mappe:</label>
                  <select value={quickFolder} onChange={(e) => setQuickFolder(e.target.value)}>
                    <option>Sendte elementer</option>
                    <option>Innboks</option>
                    <option>Utkast</option>
                    <option>Søppelpost</option>
                  </select>
                  <label>Eldre enn:</label>
                  <select value={quickYears} onChange={(e) => setQuickYears(Number(e.target.value))}>
                    <option value={1}>1 år</option>
                    <option value={2}>2 år</option>
                    <option value={3}>3 år</option>
                    <option value={5}>5 år</option>
                  </select>
                </div>
                <button className="secondary-btn" onClick={startQuickArchive} style={{ marginTop: "8px" }}>
                  {"\u{1F4E6}"} Arkiver gamle e-poster
                </button>
              </div>

              <div className="quick-section">
                <h3>{"\u{1F4C2}"} Organiser mapper</h3>
                <p className="hint">Grupper e-poster i en mappe i undermapper basert på avsender. Avsendere med flere enn terskel-antallet får egen undermappe.</p>
                <div className="quick-form">
                  <label>Mappe:</label>
                  <input
                    type="text"
                    value={orgFolder}
                    onChange={(e) => setOrgFolder(e.target.value)}
                    placeholder="F.eks. Innboks, Arkiv, Nyhetsbrev..."
                    style={{ flex: 1 }}
                  />
                  <label>Terskel:</label>
                  <select value={orgThreshold} onChange={(e) => setOrgThreshold(Number(e.target.value))}>
                    <option value={3}>3 e-poster</option>
                    <option value={5}>5 e-poster</option>
                    <option value={10}>10 e-poster</option>
                    <option value={20}>20 e-poster</option>
                  </select>
                </div>
                <button className="secondary-btn" onClick={startOrganize} style={{ marginTop: "8px" }}>
                  {"\u{1F4C2}"} Organiser nå
                </button>
              </div>
            </>
          )}
        </div>
      )}

      {phase === "working" && (
        <div className="phase-progress">
          <div className="progress-bar">
            <div
              className="progress-fill"
              style={{ width: `${progress.percent}%` }}
            />
          </div>
          <p className="progress-text">{progress.message}</p>
          <div className="spinner" />
          {results.length > 0 && (
            <div className="results" style={{ marginTop: "12px", maxHeight: "200px", overflowY: "auto" }}>
              {results.map((r, i) => (
                <p key={i}>{r}</p>
              ))}
            </div>
          )}
        </div>
      )}

      {phase === "done" && (
        <div className="phase-done">
          <h2>{"\u2705"} Rydding fullført!</h2>
          <div className="results">
            {results.map((r, i) => (
              <p key={i}>{r}</p>
            ))}
          </div>
          <button className="primary-btn" onClick={reset}>
            {"\u{1F504}"} Kjør igjen
          </button>
        </div>
      )}

      {phase === "quick-archive" && (
        <div className="phase-progress">
          <h2>{"\u{1F4E6}"} Hurtigrydding</h2>
          <p className="progress-text">{quickMsg}</p>
          <p style={{ fontSize: "24px", fontWeight: "bold", color: "#2563eb", textAlign: "center", margin: "12px 0" }}>
            {quickMoved} e-poster flyttet
          </p>
          {!quickMsg.startsWith("Ferdig") ? (
            <div className="spinner" />
          ) : (
            <button className="primary-btn" onClick={reset} style={{ marginTop: "12px" }}>
              {"\u{1F504}"} Tilbake
            </button>
          )}
        </div>
      )}

      {phase === "organize" && (
        <div className="phase-progress">
          <h2>{"\u{1F4C2}"} Organiser mapper</h2>
          <p className="progress-text">{orgMsg}</p>
          <div style={{ display: "flex", justifyContent: "center", gap: "24px", margin: "12px 0" }}>
            <div style={{ textAlign: "center" }}>
              <p style={{ fontSize: "24px", fontWeight: "bold", color: "#2563eb" }}>{orgSubs}</p>
              <p style={{ fontSize: "12px", color: "#666" }}>undermapper</p>
            </div>
            <div style={{ textAlign: "center" }}>
              <p style={{ fontSize: "24px", fontWeight: "bold", color: "#2563eb" }}>{orgMoved}</p>
              <p style={{ fontSize: "12px", color: "#666" }}>e-poster flyttet</p>
            </div>
          </div>
          {!orgMsg.startsWith("Ferdig") ? (
            <div className="spinner" />
          ) : (
            <button className="primary-btn" onClick={reset} style={{ marginTop: "12px" }}>
              {"\u{1F504}"} Tilbake
            </button>
          )}
        </div>
      )}

      {phase === "error" && (
        <div className="phase-error">
          <h2>{"\u274C"} Feil</h2>
          <p style={{ whiteSpace: "pre-wrap", wordBreak: "break-word", fontSize: "13px" }}>{error}</p>
          <button className="primary-btn" onClick={reset}>
            Prøv igjen
          </button>
        </div>
      )}
    </div>
  );
}

function typeLabel(type: string): string {
  switch (type) {
    case "create_folder":
      return "\u{1F4C1} Ny mappe";
    case "move":
      return "\u{1F4E6} Flytt";
    case "archive":
      return "\u{1F5C4}\uFE0F Arkiver";
    case "delete":
      return "\u{1F5D1}\uFE0F Slett";
    default:
      return type;
  }
}

export default App;
