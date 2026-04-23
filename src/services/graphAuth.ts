// Auth via Office Dialog API + implicit OAuth2 flow (no MSAL dependency)
// Dialog opens auth-end.html → redirects to Microsoft login → gets token → messageParent

const STORAGE_KEY = "epostrydder_session";
const TOKEN_BUFFER_MS = 5 * 60 * 1000; // refresh 5 min before expiry

let accessToken: string | null = null;
let tokenExpiry: Date | null = null;
let currentUser: string | null = null;
let debugLogger: ((msg: string) => void) | null = null;

export function setDebugLogger(fn: (msg: string) => void): void {
  debugLogger = fn;
}

function log(msg: string): void {
  console.log(`[E-postrydder] ${msg}`);
  if (debugLogger) debugLogger(msg);
}

interface SessionData {
  accessToken: string;
  expiresOn: string;
  username: string;
}

function saveSession(): void {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ accessToken, expiresOn: tokenExpiry?.toISOString(), username: currentUser })
    );
  } catch {
    // localStorage unavailable
  }
}

function isTokenValid(): boolean {
  return !!accessToken && !!tokenExpiry && tokenExpiry.getTime() - TOKEN_BUFFER_MS > Date.now();
}

async function fetchUsername(token: string): Promise<string> {
  try {
    log("Henter brukernavn fra Graph /me...");
    const res = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.ok) {
      const data = await res.json();
      const name = data.mail || data.userPrincipalName || "";
      log(`Graph /me OK: ${name}`);
      return name;
    } else {
      log(`Graph /me feilet: ${res.status} ${res.statusText}`);
    }
  } catch (e: any) {
    log(`Graph /me unntak: ${e.message}`);
  }
  return "";
}

let activeDialog: Office.Dialog | null = null;

function openAuthDialog(): Promise<string> {
  return new Promise((resolve, reject) => {
    // Force-close any previous dialog that might still be open
    if (activeDialog) {
      log("Lukker forrige dialog...");
      try { activeDialog.close(); } catch (e: any) { log(`Lukking feilet: ${e.message}`); }
      activeDialog = null;
    }

    // Check Office.context.ui availability
    if (!Office?.context?.ui?.displayDialogAsync) {
      const err = "Office.context.ui.displayDialogAsync er ikke tilgjengelig";
      log(err);
      reject(new Error(err));
      return;
    }

    const dialogUrl = `${window.location.protocol}//${window.location.host}/addin/auth-end.html`;
    log(`Åpner dialog: ${dialogUrl}`);
    log(`Office.context.requirements: ${JSON.stringify(Office?.context?.requirements || "N/A")}`);

    try {
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 60, width: 40, promptBeforeOpen: false },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            const code = asyncResult.error.code;
            const msg = asyncResult.error.message;
            log(`Dialog FEILET: kode=${code}, melding=${msg}`);
            reject(new Error(`Kunne ikke åpne innlogging (${code}): ${msg}`));
            return;
          }

          const dialog = asyncResult.value;
          activeDialog = dialog;
          log("Dialog åpnet OK — venter på svar fra innlogging...");

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
            log("Mottok melding fra dialog");
            dialog.close();
            activeDialog = null;
            try {
              const msg: string = typeof arg.message === "string" ? arg.message : arg.data;
              log(`Meldingsdata: ${msg.substring(0, 200)}`);
              const data = JSON.parse(msg);
              if (data.status === "success") {
                accessToken = data.accessToken;
                const expiresIn = parseInt(data.expiresIn || "3600", 10);
                tokenExpiry = new Date(Date.now() + expiresIn * 1000);
                log(`Token mottatt, utløper om ${expiresIn}s`);
                saveSession();
                // Fetch username in background
                fetchUsername(data.accessToken).then((name) => {
                  currentUser = name;
                  log(`Bruker: ${name}`);
                  saveSession();
                });
                resolve(data.accessToken);
              } else {
                log(`Dialog returnerte feil: ${data.error}`);
                reject(new Error(data.error || "Innlogging feilet"));
              }
            } catch (e: any) {
              log(`Parsing av dialog-svar feilet: ${e.message}`);
              reject(new Error("Ugyldig svar fra innlogging"));
            }
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
            const code = arg?.error || arg?.type || "ukjent";
            log(`Dialog-hendelse: ${code} (bruker lukket dialog?)`);
            activeDialog = null;
            reject(new Error("Innlogging ble avbrutt"));
          });
        }
      );
    } catch (e: any) {
      log(`displayDialogAsync kastet unntak: ${e.message}`);
      reject(new Error(`Dialog-feil: ${e.message}`));
    }
  });
}

export async function getGraphToken(): Promise<string> {
  if (isTokenValid()) return accessToken!;
  return openAuthDialog();
}

/** Returns the current token if valid, or throws — never opens a dialog */
export async function getGraphTokenSafe(): Promise<string> {
  if (isTokenValid()) return accessToken!;
  throw new Error("TOKEN_EXPIRED");
}

export async function signIn(): Promise<string> {
  const token = await getGraphToken();
  if (!currentUser) {
    currentUser = await fetchUsername(token);
    saveSession();
  }
  return currentUser || "";
}

export function isSignedIn(): boolean {
  return !!currentUser && isTokenValid();
}

export function getSignedInUser(): string {
  return currentUser || "";
}

export function ensureTokenValid(): boolean {
  return isTokenValid();
}

export function signOut(): void {
  accessToken = null;
  tokenExpiry = null;
  currentUser = null;
  localStorage.removeItem(STORAGE_KEY);
  log("Logget ut");
}

export async function tryRestoreSession(): Promise<boolean> {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const data: SessionData = JSON.parse(saved);
      const expiry = new Date(data.expiresOn);
      if (expiry.getTime() - TOKEN_BUFFER_MS > Date.now()) {
        accessToken = data.accessToken;
        tokenExpiry = expiry;
        currentUser = data.username;
        return true;
      }
    }
  } catch {
    // corrupt data
  }
  localStorage.removeItem(STORAGE_KEY);
  return false;
}
