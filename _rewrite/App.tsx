import React, { useState, useCallback } from "react";
import {
  hasApiKey,
  setApiKey as saveApiKey,
  getApiKey,
  analyzeMailbox,
  CleanupPlan,
} from "../services/gemini";
import {
  scanMailbox,
  moveMessages,
  createFolder,
  deleteMessages,
  MailboxSnapshot,
} from "../services/outlook";
import "./taskpane.css";

type Phase = "idle" | "scanning" | "analyzing" | "review" | "executing" | "done" | "error";

function App() {
  const [phase, setPhase] = useState<Phase>("idle");
  const [progress, setProgress] = useState({ message: "", percent: 0 });
  const [snapshot, setSnapshot] = useState<MailboxSnapshot | null>(null);
  const [plan, setPlan] = useState<CleanupPlan | null>(null);
  const [error, setError] = useState("");
  const [showSettings, setShowSettings] = useState(!hasApiKey());
  const [apiKey, setApiKeyInput] = useState(getApiKey() || "");
  const [results, setResults] = useState<string[]>([]);

  const handleSaveKey = () => {
    saveApiKey(apiKey.trim());
    setShowSettings(false);
  };

  const startScan = useCallback(async () => {
    if (!hasApiKey()) {
      setShowSettings(true);
      return;
    }
    try {
      // Phase 1: Scan mailbox
      setPhase("scanning");
      const snap = await scanMailbox((msg, pct) =>
        setProgress({ message: msg, percent: pct })
      );
      setSnapshot(snap);

      // Phase 2: AI analysis
      setPhase("analyzing");
      setProgress({ message: "Gemini analyserer postboksen din...", percent: 0 });
      const p = await analyzeMailbox(snap);
      setPlan(p);
      setPhase("review");
    } catch (e: any) {
      setError(e.message || "Ukjent feil oppstod");
      setPhase("error");
    }
  }, []);

  const toggleAction = (idx: number) => {
    if (!plan) return;
    const updated = { ...plan };
    updated.actions = updated.actions.map((a, i) =>
      i === idx ? { ...a, approved: !a.approved } : a
    );
    setPlan(updated);
  };

  const executePlan = useCallback(async () => {
    if (!plan || !snapshot) return;
    setPhase("executing");
    const log: string[] = [];
    const approved = plan.actions.filter((a) => a.approved);

    // Map existing folder names to IDs (case-insensitive)
    const folderMap: Record<string, string> = {};
    for (const f of snapshot.folders) {
      folderMap[f.name.toLowerCase()] = f.id;
    }

    for (let i = 0; i < approved.length; i++) {
      const action = approved[i];
      setProgress({
        message: action.description,
        percent: Math.round(((i + 1) / approved.length) * 100),
      });

      try {
        if (action.type === "create_folder" && action.newFolderName) {
          const id = await createFolder(action.newFolderName);
          folderMap[action.newFolderName.toLowerCase()] = id;
          log.push(`\u2705 Opprettet mappe "${action.newFolderName}"`);
        } else if (action.type === "move" && action.targetFolder) {
          let targetId = folderMap[action.targetFolder.toLowerCase()];
          if (!targetId) {
            targetId = await createFolder(action.targetFolder);
            folderMap[action.targetFolder.toLowerCase()] = targetId;
            log.push(`\u2705 Opprettet mappe "${action.targetFolder}"`);
          }
          const ids = action.messageIndices
            .map((idx) => snapshot.allMessages.find((m) => m.index === idx)?.id)
            .filter(Boolean) as string[];
          if (ids.length > 0) {
            const moved = await moveMessages(ids, targetId);
            log.push(`\u2705 Flyttet ${moved}/${ids.length} e-poster til "${action.targetFolder}"`);
          }
        } else if (action.type === "archive") {
          let archiveId = folderMap["arkiv"] || folderMap["archive"];
          if (!archiveId) {
            archiveId = await createFolder("Arkiv");
            folderMap["arkiv"] = archiveId;
            log.push(`\u2705 Opprettet mappe "Arkiv"`);
          }
          const ids = action.messageIndices
            .map((idx) => snapshot.allMessages.find((m) => m.index === idx)?.id)
            .filter(Boolean) as string[];
          if (ids.length > 0) {
            const moved = await moveMessages(ids, archiveId);
            log.push(`\u2705 Arkivert ${moved}/${ids.length} e-poster`);
          }
        } else if (action.type === "delete") {
          const ids = action.messageIndices
            .map((idx) => snapshot.allMessages.find((m) => m.index === idx)?.id)
            .filter(Boolean) as string[];
          if (ids.length > 0) {
            const deleted = await deleteMessages(ids);
            log.push(`\u2705 Slettet ${deleted}/${ids.length} e-poster`);
          }
        }
      } catch (e: any) {
        log.push(`\u274C ${action.description}: ${e.message}`);
      }
    }

    setResults(log);
    setPhase("done");
  }, [plan, snapshot]);

  const reset = () => {
    setPhase("idle");
    setSnapshot(null);
    setPlan(null);
    setError("");
    setResults([]);
    setProgress({ message: "", percent: 0 });
  };

  const approvedCount = plan ? plan.actions.filter((a) => a.approved).length : 0;

  return (
    <div className="app">
      <header>
        <h1>{"\u{1F4E7}"} E-postrydder</h1>
        <button
          className="icon-btn"
          onClick={() => setShowSettings(!showSettings)}
          title="Innstillinger"
        >
          {"\u2699\uFE0F"}
        </button>
      </header>

      {showSettings && (
        <div className="settings-panel">
          <label>Gemini API-n\u00f8kkel</label>
          <input
            type="password"
            value={apiKey}
            onChange={(e) => setApiKeyInput(e.target.value)}
            placeholder="AIza..."
          />
          <button onClick={handleSaveKey} disabled={!apiKey.trim()}>
            Lagre
          </button>
          <p className="hint">
            Hent n\u00f8kkel fra{" "}
            <a href="https://aistudio.google.com/apikey" target="_blank" rel="noopener noreferrer">
              Google AI Studio
            </a>
          </p>
        </div>
      )}

      {phase === "idle" && (
        <div className="phase-idle">
          <div className="intro-card">
            <p>
              E-postrydder skanner <strong>alle mappene</strong> i postboksen din,
              lar Gemini AI analysere innholdet, og foresl\u00e5r en ryddeplan du kan godkjenne f\u00f8r noe endres.
            </p>
            <ul>
              <li>{"\u{1F50D}"} Skanner alle mapper og e-poster</li>
              <li>{"\u{1F9E0}"} AI analyserer m\u00f8nstre og rot</li>
              <li>{"\u2705"} Du godkjenner hver handling</li>
              <li>{"\u{1F4E6}"} Flytter, arkiverer og rydder</li>
            </ul>
          </div>
          <button className="primary-btn" onClick={startScan} disabled={!hasApiKey()}>
            {"\u{1F50D}"} Start skanning
          </button>
          {!hasApiKey() && (
            <p className="warning">Legg inn Gemini API-n\u00f8kkel f\u00f8rst {"\u2B06\uFE0F"}</p>
          )}
        </div>
      )}

      {(phase === "scanning" || phase === "analyzing" || phase === "executing") && (
        <div className="phase-progress">
          <div className="progress-bar">
            <div
              className="progress-fill"
              style={{ width: `${phase === "analyzing" ? 100 : progress.percent}%` }}
            />
          </div>
          <p className="progress-text">{progress.message}</p>
          {phase === "analyzing" && <div className="spinner" />}
        </div>
      )}

      {phase === "review" && plan && (
        <div className="phase-review">
          <div className="summary-card">
            <h2>Analyse</h2>
            <p>{plan.summary}</p>
            {plan.observations?.length > 0 && (
              <ul>
                {plan.observations.map((o, i) => (
                  <li key={i}>{o}</li>
                ))}
              </ul>
            )}
            <div className="stats">
              {"\u{1F4C1}"} {snapshot?.folders.length} mapper {"\u00B7"}{" "}
              {"\u{1F4EC}"} {snapshot?.totalMessages} e-poster {"\u00B7"}{" "}
              {"\u{1F4ED}"} {snapshot?.totalUnread} uleste
            </div>
          </div>

          <h2>
            Ryddeplan ({approvedCount}/{plan.actions.length} valgt)
          </h2>
          {plan.actions.map((action, i) => (
            <div
              key={i}
              className={`action-card ${action.approved ? "approved" : "rejected"}`}
            >
              <label className="action-header">
                <input
                  type="checkbox"
                  checked={action.approved}
                  onChange={() => toggleAction(i)}
                />
                <span className="action-type">{typeLabel(action.type)}</span>
                <span className="action-desc">{action.description}</span>
              </label>
              <p className="action-reason">{action.reason}</p>
              {action.messageIndices?.length > 0 && (
                <p className="action-count">
                  {action.messageIndices.length} e-post
                  {action.messageIndices.length !== 1 ? "er" : ""}
                </p>
              )}
            </div>
          ))}

          <div className="actions-bar">
            <button className="secondary-btn" onClick={reset}>
              Avbryt
            </button>
            <button
              className="primary-btn"
              onClick={executePlan}
              disabled={approvedCount === 0}
            >
              {"\u2705"} Utf\u00f8r rydding
            </button>
          </div>
        </div>
      )}

      {phase === "done" && (
        <div className="phase-done">
          <h2>{"\u2705"} Rydding fullf\u00f8rt!</h2>
          <div className="results">
            {results.map((r, i) => (
              <p key={i}>{r}</p>
            ))}
          </div>
          <button className="primary-btn" onClick={reset}>
            {"\u{1F504}"} Kj\u00f8r ny analyse
          </button>
        </div>
      )}

      {phase === "error" && (
        <div className="phase-error">
          <h2>{"\u274C"} Feil</h2>
          <p>{error}</p>
          <button className="primary-btn" onClick={reset}>
            Pr\u00f8v igjen
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
