"use client";

import { CopilotKit, useCopilotAction, useCopilotAdditionalInstructions } from "@copilotkit/react-core";
import { CopilotSidebar } from "@copilotkit/react-ui";
import "@copilotkit/react-ui/styles.css";
import { FileSignature, Search, ListChecks } from "lucide-react";
import { createContext, useContext, useRef, useState, useEffect, useCallback } from "react";
import { BoxPreviewPanel, type PreviewData, type RequestSummary } from "./BoxPreviewPanel";

const AGENT_INSTRUCTIONS = `You are the Box Sign AI Assistant. You help users create and manage e-signature requests in Box.

## Scope: document signing only
- You must **only** answer questions and perform actions related to **document signing with Box Sign**: searching files in Box, creating signature requests, listing or checking status of sign requests, cancelling or resending requests, and explaining how Box Sign options work (expiration, reminders, signer order, email subject/message).
- For **any other topic** (weather, news, general knowledge, other products, off-topic chat), do **not** answer or use tools. Reply briefly and redirect: "I can only help with Box Sign—searching documents, creating and managing signature requests. Is there something you'd like to do with a document to sign?" Do not provide information about weather, locations, or anything unrelated to the signing process.

## CRITICAL: Preview and confirm before sending
- **Always** call **prepare_signature_request** first (with the same parameters you would use for create_signature_request). This shows the user a document preview and request details next to the chat. Tell the user: "I've prepared the request. You can see the document preview and details. Is this the correct document to sign? Reply **Yes** or **Confirm** to send the signature request."
- **Only after** the user explicitly confirms (e.g. "yes", "confirm", "looks good", "send it", "go ahead") call **create_signature_request** with the same parameters you used in prepare_signature_request. Never call create_signature_request without having first called prepare_signature_request and received user confirmation.

## CRITICAL: Required data before preparing/creating
- Never call prepare_signature_request or create_signature_request until you have **both**: (1) a valid **file ID** (from search_files or explicitly from the user), (2) at least one **real participant email** (signer, approver, or final_copy_reader) that the user stated (never use example@example.com or placeholders—the API rejects them).
- If the user gives participants + options (expiration, reminders, subject, etc.) but **no document**, reply: "I have the participant(s) and options. Which document should I send for signature? You can give me a file ID or I can search your Box (e.g. by file name)."
- If the user gives only a file ID and no participants, ask: "Who should be involved? Please provide at least one signer (and optionally approvers or people who should get a copy when done)."

## Handling combined / single-message requests
Users often give many details in one message. **Parse and extract** everything they said, then ask only for what is missing (usually the document or the signer).

**Example:** "One signer, alice@company.com, valid for a week, please enable reminder one day before expiration, also add subject '[NDA] Important document from Box'"
- Extract: participants = [{ email: "alice@company.com", role: "signer" }], daysValid = 7, areRemindersEnabled = true, emailSubject = "[NDA] Important document from Box".
- **Reminders:** Box Sign uses a fixed cadence (days 3, 8, 13, 18), but only reminders that occur before expiration are relevant. If daysValid is 7, only day 3 applies. Set areRemindersEnabled = true and tell the user the effective reminder days for their expiration window.
- Document is missing → ask: "Which document should I use? Give me a file ID or I can search Box by name."

**Example:** "Create a sign request for file 123456 with john@company.com and jane@company.com, sequential, expire in 30 days, subject 'Contract to sign'"
- Extract: fileId = "123456", participants = [{ email: "john@company.com", role: "signer" }, { email: "jane@company.com", role: "signer" }], isSequential = true, daysValid = 30, emailSubject = "Contract to sign". Call create_signature_request with these participants.

**Example (with approver):** "Send the NDA for signature to alice@co.com, and legal@co.com should approve it; expire in 14 days"
- Extract: need file (search or ask), participants = [{ email: "alice@co.com", role: "signer" }, { email: "legal@co.com", role: "approver" }], daysValid = 14. Do **not** put legal@co.com as a signer—they are an approver.

## Box Sign roles (CRITICAL: do not treat approvers as signers)
Box Sign has three distinct participant roles. **Never** assign someone who is described as an approver or "get a copy" as a signer.

- **signer**: Must sign the document (place signature). Use when the user says "sign", "signer", "needs to sign", "signature", "pass it to [name] to sign".
- **approver**: Reviews the document and approves or declines; does **not** sign. Use when the user says "approver", "approve", "approval", "review and approve", "one approver:", "for approval".
- **final_copy_reader**: Only receives a copy of the final signed document when done; no action required. Use when the user says "get a copy", "receive a copy", "final copy", "cc", "for their records".

**RULE for approver + signer workflows:** When the user says "one approver" and "one signer" (or "approver then signer", "pass it to a signer"), you MUST call create_signature_request with **approverEmails** and **signerEmails** as separate arrays. Do NOT rely only on participants for this case. Example: approverEmails: ["legal@company.com"], signerEmails: ["manager@company.com"], isSequential: true, plus fileId, parentFolderId (from search if needed), daysValid, areRemindersEnabled, emailSubject. The backend will apply correct roles and order (approvers first, then signers, then final copy readers).

**When to use participants:** Use the **participants** array when you have a custom order or mixed roles and are sure each object has the correct role: [{ email: string, role: "signer"|"approver"|"final_copy_reader" }]. Order in the array = sequence when isSequential is true. Example: "approver legal@company.com then signer manager@company.com" → either use approverEmails + signerEmails (preferred), or participants: [{ email: "legal@company.com", role: "approver" }, { email: "manager@company.com", role: "signer" }], isSequential: true.

**When only signers:** If the user only mentions "signers" (everyone signs), use signerEmails: ["a@b.com", "c@d.com"] or participants: [{ email: "a@b.com", role: "signer" }, { email: "c@d.com", role: "signer" }].

## Natural language → API parameters (use this mapping)
- **Expiration:** "valid for a week" / "one week" / "7 days" → daysValid: 7. "Two weeks" → 14. "One month" → 30. "90 days" → 90. "No expiration" / "don't expire" → 0. Max 730.
- **Reminders:** "enable reminder" / "remind them" / "send reminders" / "reminder one day before expiration" → areRemindersEnabled: true. Box uses fixed days 3, 8, 13, 18; communicate only the effective days within \`daysValid\` (e.g. 7 days -> day 3 only).
- **Email subject:** "subject [...]" / "add subject '...'" / "email subject ..." → emailSubject: use the exact string they want (e.g. "[NDA] Important document from Box").
- **Email message:** "message body" / "email message" / "custom message in the email" → emailMessage (supports basic HTML).
- **Signing order:** "sequential" / "in order" / "John first then Sarah" → isSequential: true; put participants in the order they specified.
- **Request name:** "name this request" / "call it 'Q4 NDA'" → name.
- At least one participant is required. Prefer **approverEmails/signerEmails/finalCopyReaderEmails** for role-based workflows. Use **participants** only when you need explicit custom ordering.

## Two user paths
1. **Simple:** User gives document + signer(s) only (no expiration, reminders, subject, message, or other options). **Before calling prepare_signature_request**, suggest additional customization options they haven't specified:
   - "I can create a sign request with those details. Before I prepare it, would you like to customize any of these options?"
   - List options they haven't mentioned: expiration date (e.g., 7, 14, or 30 days), email reminders, custom email subject or message, additional roles (approvers/final copy recipients), signing order (if multiple signers), or request name.
   - Ask: "Would you like me to suggest these customization options for future requests, or would you prefer I proceed with defaults unless you specify otherwise?"
   - If user says they don't want suggestions, note their preference and skip this step in future interactions during this session. If they provide options or say "use defaults" / "proceed", call prepare_signature_request; after they confirm the preview, call create_signature_request.
2. **Extended / combined:** User gives document + signer(s) + some options in one go. Use everything they said (map via the table above). If only document is missing, ask for document. If only signers are missing, ask for signers. Then call prepare_signature_request; after user confirms, call create_signature_request.

## Bulk operations

### Bulk creation (multiple requests at once)
- **When to use bulk operations:** User explicitly requests sending the SAME document to MULTIPLE different recipients, OR sending MULTIPLE different documents. Examples: "Send this to 5 people", "Create 3 sign requests for these files".
- **When NOT to use bulk:** Single document with multiple signers/approvers on one request (use regular create_signature_request instead).
- **Bulk workflow (must follow this order):**
  1. **First call ONLY**: **bulk_prepare_signature_requests** with the requests array. This shows the user all documents in a preview panel they can browse through. Tell the user how many requests you've prepared and that they can review all documents.
  2. **Wait for user confirmation**: User must reply "yes", "confirm", "looks good", etc.
  3. **Then call ONLY ONCE**: **bulk_create_signature_requests** with the same requests array. This creates all requests.
  4. **Stop and report**: After bulk_create_signature_requests returns the summary, you are DONE. Report the summary to the user. Do NOT call any other actions (no status checks, no verification, no additional calls). The bulk operation is complete.
- The requests parameter is an array where each item has: fileId (required), participants or approverEmails/signerEmails (required), and optional parameters.
- Example use cases:
  - Same document to multiple recipients separately: requests array with multiple objects having the same fileId but different signerEmails (each recipient gets their own request)
  - Different documents to different recipients: requests array with objects having different fileIds and different signerEmails
- The user can navigate through all documents using Previous/Next buttons in the preview panel.
- Return the summary from bulk_create_signature_requests showing how many succeeded/failed.

### Bulk cancellation
- If user asks to remove/cancel/delete **all** sign requests matching a phrase (e.g. "Remove all sign requests for Vendor agreement"), call **bulk_cancel_signature_requests** with that phrase.
- Use case-insensitive partial matching against request name and ID.
- Report how many were cancelled, skipped, and not found.

## After creating or when returning details (single requests only)
This section applies to SINGLE signature requests created with create_signature_request. For bulk operations, the summary returned by bulk_create_signature_requests is final - do not call additional actions.

For single requests, always include participant roles in the response. Use heading **Participants (role and status)**, not "Signers", when roles are mixed.
- Show each participant as: email — role — status (e.g. "alice@company.com — Approver — pending").
- If role is unavailable from API, label as "Participant" rather than assuming signer.
- Include Request ID, current status, expiration info (if present), and Prepare URL (if present).
- Never invent details from memory. If user asks for request details or status, call **get_signature_request_status** with a request ID (use the most recently created request ID if the user says "this sign request").
- If user says "return details of this sign request", "details of the request I just created", or "status of this request", prefer **get_latest_signature_request_status** (no parameters) so details always come from the API.

## When creation fails: explain why and what to do
If create_signature_request returns an error, **tell the user clearly why the sign request was not created** and what they can do next.

- **File in root folder / parentFolderId required:** If the error says the file is in the root folder or asks for a parentFolderId, explain in plain language: "The sign request wasn't created because this document is in your Box root folder. Box Sign requires the file to be inside a folder. To fix this: (1) I can search for the file—search results include a parentFolderId you can use—then create the request again with that parentFolderId, or (2) Create a folder in Box, move or upload the file there, and use that folder's ID when we create the request." Offer to run search_files so the user gets a parentFolderId from the results.
- **File not found (404):** Tell the user the file ID may be wrong or the file was moved/deleted; suggest searching Box again or confirming the ID.
- **Placeholder emails rejected:** Remind them to use real signer email addresses, not example@example.com.
- For any other error, relay the error message and suggest checking their Box account or trying again.`;

function handleCopilotError(errorEvent: {
  type?: string;
  error?: { message?: string; status?: number; requestId?: string };
  context?: { response?: { status?: number; statusText?: string }; metadata?: { requestId?: string } };
}) {
  const err = errorEvent.error;
  const ctx = errorEvent.context;
  const status = err?.status ?? ctx?.response?.status;
  const requestId = err?.requestId ?? ctx?.metadata?.requestId;
  const msg =
    err?.message ??
    ctx?.response?.statusText ??
    (status === 400
      ? "Bad request — check that /api/copilotkit is available and OPENAI_API_KEY is set."
      : "Something went wrong.");
  const detail = `${msg}${requestId ? ` (Request ID: ${requestId})` : ""}`;
  console.error("[CopilotKit]", detail, errorEvent);
  if (typeof window !== "undefined" && window.alert) {
    window.alert(`Error: ${detail}`);
  }
}

/** Single backend for all Box actions. CopilotKit orchestrates; this route runs Box SDK server-side. */
async function boxAction<T = Record<string, unknown>>(
  action: string,
  params: Record<string, unknown> = {}
): Promise<T & { error?: string; debug?: { hint?: string } }> {
  const res = await fetch("/api/box", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action, ...params }),
  });
  const text = await res.text();
  if (!text) {
    if (!res.ok)
      return {
        error: res.status === 400 ? "Bad request." : res.statusText || `Error ${res.status}`,
      } as T & { error?: string };
    return {} as T & { error?: string };
  }
  if (res.status === 404 || (text.trimStart().startsWith("<!") && !res.ok)) {
    return {
      error: res.status === 404
        ? "Service unavailable (404). The Box API route may be missing. Restart the dev server."
        : `Request failed (${res.status}). Server may have returned an error page.`,
    } as T & { error?: string };
  }
  try {
    return JSON.parse(text) as T & { error?: string; debug?: { hint?: string } };
  } catch {
    return {
      error: res.ok ? "Invalid response." : text || res.statusText || `Error ${res.status}`,
    } as T & { error?: string };
  }
}

function chatError(
  data: { error?: string; debug?: { hint?: string; userMessage?: string } } | null,
  fallback: string
): string {
  const msg = data?.error?.trim() || fallback;
  const userMessage = data?.debug?.userMessage?.trim();
  if (userMessage) return userMessage;
  const hint = data?.debug?.hint?.trim();
  return hint ? `${msg} ${hint}` : msg;
}

function roleLabel(role: string | undefined): string {
  switch (role) {
    case "approver":
      return "Approver";
    case "final_copy_reader":
      return "Final copy reader";
    case "signer":
      return "Signer";
    default:
      return "Participant";
  }
}

type ActiveSignRequest = {
  id: string;
  name?: string;
  status: string;
  signers: Array<{ email?: string; role?: string; status?: string }>;
  autoExpireAt?: string;
  createdAt?: string;
};

const INACTIVE_SIGNATURE_STATUSES = new Set([
  "cancelled",
  "declined",
  "expired",
  "completed",
]);

function isActiveSignRequestStatus(status?: string): boolean {
  return !INACTIVE_SIGNATURE_STATUSES.has((status ?? "").toLowerCase());
}

type PreviewContextValue = {
  previewData: PreviewData | null;
  setPreviewData: (data: PreviewData | null) => void;
  activeSignRequests: ActiveSignRequest[];
  setActiveSignRequests: (requests: ActiveSignRequest[]) => void;
  activeSignRequestsLoading: boolean;
  setActiveSignRequestsLoading: (isLoading: boolean) => void;
};
const PreviewContext = createContext<PreviewContextValue | null>(null);

export default function BoxSignAssistant() {
  const [previewData, setPreviewData] = useState<PreviewData | null>(null);
  const [activeSignRequests, setActiveSignRequests] = useState<ActiveSignRequest[]>([]);
  const [activeSignRequestsLoading, setActiveSignRequestsLoading] = useState(false);

  return (
    <CopilotKit
      runtimeUrl="/api/copilotkit"
      onError={handleCopilotError}
      showDevConsole={process.env.NODE_ENV === "development"}
    >
      <PreviewContext.Provider
        value={
          {
            previewData,
            setPreviewData,
            activeSignRequests,
            setActiveSignRequests,
            activeSignRequestsLoading,
            setActiveSignRequestsLoading,
          } as PreviewContextValue
        }
      >
        <BoxSignActions />
        <CopilotSidebar
          defaultOpen
          clickOutsideToClose={false}
          labels={{
            title: "Box Sign AI Assistant",
            initial:
              "Get help with signing documents in Box, creating and managing signature workflow requests. \n\n**Simple flow**: Provide a document name or ID and who should sign — I'll show you a preview and request details, then you confirm to send.\n\n**Full workflow**: set expiration date, automatic reminders, signing order, or custom email subject and message",
          }}
        />
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            flex: 1,
            minHeight: "100vh",
            marginRight: "440px",
          }}
        >
          {previewData ? (
            <div
              style={{
                flex: 1,
                minHeight: "calc(100vh - 2rem)",
                display: "flex",
                width: "100%",
                maxWidth: "none",
                margin: "1rem auto",
                padding: "0 1rem",
                boxSizing: "border-box",
              }}
            >
              <BoxPreviewPanel
                data={previewData}
                onDismiss={() => setPreviewData(null)}
              />
            </div>
          ) : (
            <main
              style={{
                padding: "var(--space-2xl) var(--space-xl)",
                maxWidth: "72rem",
                margin: "0 auto",
                width: "100%",
              }}
            >
        <div
          style={{
            background: "linear-gradient(135deg, #0061D5 0%, #003D8F 100%)",
            borderRadius: "var(--radius-xl)",
            padding: "var(--space-2xl)",
            color: "white",
            marginBottom: "var(--space-2xl)",
            boxShadow: "var(--shadow-lg)",
          }}
        >
          <h1
            style={{
              margin: 0,
              fontSize: "var(--text-3xl)",
              fontWeight: "var(--font-bold)",
              display: "flex",
              alignItems: "center",
              gap: "var(--space-md)",
              letterSpacing: "-0.025em",
            }}
          >
            <FileSignature size={32} strokeWidth={2.5} />
            Box Sign AI Assistant
          </h1>
          <p
            style={{
              margin: "var(--space-md) 0 0",
              opacity: 0.95,
              fontSize: "var(--text-lg)",
              maxWidth: "42rem",
              lineHeight: 1.6,
            }}
          >
            Create and manage e-signature requests using natural language.
          </p>
        </div>
        
        <div
          style={{
            background: "white",
            border: "var(--border-default)",
            borderRadius: "var(--radius-lg)",
            padding: "var(--space-xl)",
            marginBottom: "var(--space-xl)",
            boxShadow: "var(--shadow-sm)",
          }}
        >
          <h2
            style={{
              margin: 0,
              fontSize: "var(--text-xl)",
              fontWeight: "var(--font-semibold)",
              color: "var(--neutral-900)",
              marginBottom: "var(--space-md)",
            }}
          >
            Getting Started
          </h2>
          <p
            style={{
              margin: 0,
              color: "var(--neutral-600)",
              fontSize: "var(--text-base)",
              lineHeight: 1.7,
              marginBottom: "var(--space-lg)",
            }}
          >
            Open the chat panel on the right to get started. You can search for
            documents in Box, create signature requests with signer emails, set
            sequential signing, reminders, and expiration.
          </p>
          
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
              gap: "var(--space-md)",
            }}
          >
            <div
              style={{
                padding: "var(--space-lg)",
                background: "var(--neutral-50)",
                borderRadius: "var(--radius-md)",
                border: "var(--border-subtle)",
              }}
            >
              <div
                style={{
                  width: "40px",
                  height: "40px",
                  borderRadius: "var(--radius-md)",
                  background: "linear-gradient(135deg, #0061D5 0%, #003D8F 100%)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  marginBottom: "var(--space-md)",
                }}
              >
                <Search size={20} color="white" strokeWidth={2.5} />
              </div>
              <h3
                style={{
                  margin: 0,
                  fontSize: "var(--text-base)",
                  fontWeight: "var(--font-semibold)",
                  color: "var(--neutral-900)",
                  marginBottom: "var(--space-xs)",
                }}
              >
                Search Documents
              </h3>
              <p
                style={{
                  margin: 0,
                  fontSize: "var(--text-sm)",
                  color: "var(--neutral-600)",
                  lineHeight: 1.5,
                }}
              >
                Find files in your Box account by name or description
              </p>
            </div>
            
            <div
              style={{
                padding: "var(--space-lg)",
                background: "var(--neutral-50)",
                borderRadius: "var(--radius-md)",
                border: "var(--border-subtle)",
              }}
            >
              <div
                style={{
                  width: "40px",
                  height: "40px",
                  borderRadius: "var(--radius-md)",
                  background: "linear-gradient(135deg, #0061D5 0%, #003D8F 100%)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  marginBottom: "var(--space-md)",
                }}
              >
                <FileSignature size={20} color="white" strokeWidth={2.5} />
              </div>
              <h3
                style={{
                  margin: 0,
                  fontSize: "var(--text-base)",
                  fontWeight: "var(--font-semibold)",
                  color: "var(--neutral-900)",
                  marginBottom: "var(--space-xs)",
                }}
              >
                Create Requests
              </h3>
              <p
                style={{
                  margin: 0,
                  fontSize: "var(--text-sm)",
                  color: "var(--neutral-600)",
                  lineHeight: 1.5,
                }}
              >
                Set up signature workflows with multiple participants
              </p>
            </div>
            
            <div
              style={{
                padding: "var(--space-lg)",
                background: "var(--neutral-50)",
                borderRadius: "var(--radius-md)",
                border: "var(--border-subtle)",
              }}
            >
              <div
                style={{
                  width: "40px",
                  height: "40px",
                  borderRadius: "var(--radius-md)",
                  background: "linear-gradient(135deg, #0061D5 0%, #003D8F 100%)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  marginBottom: "var(--space-md)",
                }}
              >
                <ListChecks size={20} color="white" strokeWidth={2.5} />
              </div>
              <h3
                style={{
                  margin: 0,
                  fontSize: "var(--text-base)",
                  fontWeight: "var(--font-semibold)",
                  color: "var(--neutral-900)",
                  marginBottom: "var(--space-xs)",
                }}
              >
                Track Progress
              </h3>
              <p
                style={{
                  margin: 0,
                  fontSize: "var(--text-sm)",
                  color: "var(--neutral-600)",
                  lineHeight: 1.5,
                }}
              >
                Monitor active requests and participant status
              </p>
            </div>
          </div>
        </div>
      </main>
          )}
          <ActiveSigningRequestsPanel
            requests={activeSignRequests}
            loading={activeSignRequestsLoading}
          />
        </div>
      </PreviewContext.Provider>
    </CopilotKit>
  );
}

function QuickAction({
  icon,
  label,
}: {
  icon: React.ReactNode;
  label: string;
}) {
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        gap: "0.5rem",
        padding: "0.75rem 1rem",
        background: "var(--background)",
        border: "1px solid rgba(0,0,0,0.08)",
        borderRadius: "12px",
        fontSize: "0.9rem",
      }}
    >
      {icon}
      <span>{label}</span>
    </div>
  );
}

function buildParticipantsFromParams({
  participantsRaw,
  approverEmailsRaw,
  signerEmailsRaw,
  finalCopyReaderEmailsRaw,
}: {
  participantsRaw: unknown;
  approverEmailsRaw: unknown;
  signerEmailsRaw: unknown;
  finalCopyReaderEmailsRaw: unknown;
}): Array<{ email: string; role: string }> {
  let rawList: Array<{ email?: string; role?: string }> = [];
  if (Array.isArray(participantsRaw)) {
    rawList = participantsRaw as Array<{ email?: string; role?: string }>;
  } else if (participantsRaw && typeof participantsRaw === "object") {
    const obj = participantsRaw as Record<string, unknown>;
    if (typeof (obj as { email?: string }).email === "string") {
      rawList = [participantsRaw as { email?: string; role?: string }];
    } else {
      rawList = Object.keys(obj)
        .filter((k) => /^\d+$/.test(k))
        .sort((a, b) => Number(a) - Number(b))
        .map((k) => obj[k] as { email?: string; role?: string })
        .filter((p) => p && typeof p === "object");
    }
  } else if (typeof participantsRaw === "string") {
    try {
      const parsed = JSON.parse(participantsRaw) as unknown;
      rawList = Array.isArray(parsed) ? (parsed as Array<{ email?: string; role?: string }>) : parsed && typeof parsed === "object" && (parsed as { email?: string }).email ? [parsed as { email?: string; role?: string }] : [];
    } catch {
      rawList = [];
    }
  }
  const approverEmails = Array.isArray(approverEmailsRaw) ? (approverEmailsRaw as string[]).map((e) => String(e).trim()).filter(Boolean) : approverEmailsRaw != null ? [String(approverEmailsRaw).trim()].filter(Boolean) : [];
  const signerEmails = Array.isArray(signerEmailsRaw) ? (signerEmailsRaw as string[]).map((e) => String(e).trim()).filter(Boolean) : signerEmailsRaw != null ? [String(signerEmailsRaw).trim()].filter(Boolean) : [];
  const finalCopyReaderEmails = Array.isArray(finalCopyReaderEmailsRaw) ? (finalCopyReaderEmailsRaw as string[]).map((e) => String(e).trim()).filter(Boolean) : finalCopyReaderEmailsRaw != null ? [String(finalCopyReaderEmailsRaw).trim()].filter(Boolean) : [];
  const hasRoleLists = approverEmails.length > 0 || signerEmails.length > 0 || finalCopyReaderEmails.length > 0;
  if (hasRoleLists) {
    return [
      ...approverEmails.map((email) => ({ email, role: "approver" })),
      ...signerEmails.map((email) => ({ email, role: "signer" })),
      ...finalCopyReaderEmails.map((email) => ({ email, role: "final_copy_reader" })),
    ];
  }
  const getStr = (p: { email?: string; role?: string }) => (x: string): string =>
    ((p as Record<string, unknown>)[x] ?? "").toString().trim();
  return rawList
    .map((p: { email?: string; role?: string }) => {
      const g = getStr(p);
      const email = (g("email") || g("Email")).trim();
      if (!email) return null;
      const role = (g("role") || g("Role") || "signer").toLowerCase();
      return { email, role: ["signer", "approver", "final_copy_reader"].includes(role) ? role : "signer" };
    })
    .filter((p): p is { email: string; role: string } => p != null);
}

function ActiveSigningRequestsPanel({
  requests,
  loading,
}: {
  requests: ActiveSignRequest[];
  loading: boolean;
}) {
  const statusColor = (status: string) => {
    const s = status.toLowerCase();
    if (s === "signed" || s === "completed") return "#0b8a4b";
    if (s === "declined" || s === "cancelled" || s === "expired") return "#b42318";
    return "#1d4ed8";
  };
  const formatDate = (value?: string) => {
    if (!value) return "—";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "—";
    return date.toLocaleString();
  };

  return (
    <section
      style={{
        margin: "0 auto var(--space-xl)",
        padding: "var(--space-xl)",
        background: "white",
        border: "var(--border-default)",
        borderRadius: "var(--radius-lg)",
        boxShadow: "var(--shadow-sm)",
        maxWidth: "64rem",
        width: "calc(100% - 4rem)",
      }}
    >
      <h2
        style={{
          margin: 0,
          fontSize: "var(--text-xl)",
          fontWeight: "var(--font-semibold)",
          display: "flex",
          alignItems: "center",
          gap: "var(--space-sm)",
          color: "var(--neutral-900)",
          letterSpacing: "-0.0125em",
        }}
      >
        <ListChecks size={22} strokeWidth={2.5} />
        Active Signing Requests
      </h2>
      <p
        style={{
          margin: "var(--space-sm) 0 var(--space-lg)",
          fontSize: "var(--text-sm)",
          color: "var(--neutral-600)",
          lineHeight: 1.5,
        }}
      >
        {loading
          ? "Refreshing active requests..."
          : "This list auto-refreshes after creating a sign request."}
      </p>
      {loading && !requests.length ? (
        <div
          style={{
            padding: "var(--space-2xl)",
            textAlign: "center",
            color: "var(--neutral-500)",
          }}
        >
          <p style={{ margin: 0, fontSize: "var(--text-sm)" }}>Loading requests...</p>
        </div>
      ) : requests.length === 0 ? (
        <div
          style={{
            padding: "var(--space-2xl)",
            textAlign: "center",
            background: "var(--neutral-50)",
            borderRadius: "var(--radius-md)",
            border: "var(--border-subtle)",
          }}
        >
          <div
            style={{
              width: "56px",
              height: "56px",
              borderRadius: "var(--radius-lg)",
              background: "var(--neutral-100)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              margin: "0 auto var(--space-md)",
            }}
          >
            <ListChecks size={28} color="var(--neutral-400)" strokeWidth={2} />
          </div>
          <p
            style={{
              margin: 0,
              fontSize: "var(--text-base)",
              fontWeight: "var(--font-medium)",
              color: "var(--neutral-700)",
              marginBottom: "var(--space-xs)",
            }}
          >
            No active signing requests
          </p>
          <p
            style={{
              margin: 0,
              fontSize: "var(--text-sm)",
              color: "var(--neutral-500)",
            }}
          >
            Create your first signature request using the chat panel
          </p>
        </div>
      ) : (
        <div
          style={{
            border: "var(--border-default)",
            borderRadius: "var(--radius-md)",
            overflow: "hidden",
            background: "white",
          }}
        >
          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              minWidth: 760,
              fontSize: "var(--text-sm)",
            }}
          >
            <thead
              style={{
                background: "var(--neutral-50)",
                textAlign: "left",
              }}
            >
              <tr>
                <th style={{ padding: "var(--space-md) var(--space-md)", fontWeight: "var(--font-semibold)", fontSize: "var(--text-xs)", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--neutral-600)" }}>Request</th>
                <th style={{ padding: "var(--space-md) var(--space-md)", fontWeight: "var(--font-semibold)", fontSize: "var(--text-xs)", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--neutral-600)" }}>Status</th>
                <th style={{ padding: "var(--space-md) var(--space-md)", fontWeight: "var(--font-semibold)", fontSize: "var(--text-xs)", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--neutral-600)" }}>Participants</th>
                <th style={{ padding: "var(--space-md) var(--space-md)", fontWeight: "var(--font-semibold)", fontSize: "var(--text-xs)", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--neutral-600)" }}>Expires</th>
                <th style={{ padding: "var(--space-md) var(--space-md)", fontWeight: "var(--font-semibold)", fontSize: "var(--text-xs)", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--neutral-600)" }}>Created</th>
              </tr>
            </thead>
            <tbody>
              {requests.map((request) => (
                <tr 
                  key={request.id} 
                  style={{ 
                    borderTop: "var(--border-subtle)",
                    transition: "background-color 150ms ease",
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = "var(--neutral-50)";
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = "transparent";
                  }}
                >
                  <td style={{ padding: "var(--space-md)", verticalAlign: "top" }}>
                    <div style={{ fontWeight: "var(--font-semibold)", fontSize: "var(--text-base)", marginBottom: "var(--space-xs)", color: "var(--neutral-900)" }}>{request.name || "Untitled request"}</div>
                    <div style={{ color: "var(--neutral-500)", fontSize: "var(--text-xs)", fontFamily: "monospace" }}>
                      {request.id}
                    </div>
                  </td>
                  <td style={{ padding: "var(--space-md)", verticalAlign: "top" }}>
                    <span
                      style={{
                        display: "inline-block",
                        padding: "0.375rem 0.75rem",
                        borderRadius: "var(--radius-sm)",
                        fontWeight: "var(--font-semibold)",
                        fontSize: "var(--text-xs)",
                        color: statusColor(request.status),
                        background: `${statusColor(request.status)}15`,
                        border: `1px solid ${statusColor(request.status)}30`,
                      }}
                    >
                      {request.status}
                    </span>
                  </td>
                  <td style={{ padding: "var(--space-md)", verticalAlign: "top" }}>
                    {(request.signers || []).length ? (
                      <div style={{ display: "flex", flexDirection: "column", gap: "var(--space-xs)" }}>
                        {request.signers.slice(0, 3).map((signer, index) => (
                          <div key={`${request.id}-${index}`} style={{ fontSize: "var(--text-sm)", color: "var(--neutral-700)" }}>
                            <span style={{ fontWeight: "var(--font-medium)" }}>{signer.email || "Unknown"}</span>
                            <span style={{ color: "var(--neutral-500)", marginLeft: "var(--space-xs)" }}>({roleLabel(signer.role || "signer")})</span>
                          </div>
                        ))}
                        {request.signers.length > 3 ? (
                          <span style={{ color: "var(--neutral-500)", fontSize: "var(--text-xs)" }}>+{request.signers.length - 3} more</span>
                        ) : null}
                      </div>
                    ) : (
                      <span style={{ color: "var(--neutral-400)" }}>—</span>
                    )}
                  </td>
                  <td style={{ padding: "var(--space-md)", verticalAlign: "top", fontSize: "var(--text-sm)", color: "var(--neutral-700)" }}>
                    {formatDate(request.autoExpireAt)}
                  </td>
                  <td style={{ padding: "var(--space-md)", verticalAlign: "top", fontSize: "var(--text-sm)", color: "var(--neutral-700)" }}>
                    {formatDate(request.createdAt)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}

function BoxSignActions() {
  useCopilotAdditionalInstructions({ instructions: AGENT_INSTRUCTIONS });
  const previewContext = useContext(PreviewContext);
  const setPreviewData = previewContext?.setPreviewData ?? (() => {});
  const setActiveSignRequests = previewContext?.setActiveSignRequests ?? (() => {});
  const setActiveSignRequestsLoading =
    previewContext?.setActiveSignRequestsLoading ?? (() => {});
  const pendingCreateSignatureRef = useRef<string | null>(null);
  const lastCreatedSignRequestIdRef = useRef<string | null>(null);
  const signRequestRoleCacheRef = useRef<
    Map<string, Array<{ email: string; role: string; order?: number }>>
  >(new Map());

  const normalizeEmail = (email?: string) => (email ?? "").trim().toLowerCase();
  const resolveRole = (
    signRequestId: string | undefined,
    participant: { email?: string; role?: string; order?: number },
    index: number
  ): string | undefined => {
    if (participant.role) return participant.role;
    if (!signRequestId) return undefined;
    const cached = signRequestRoleCacheRef.current.get(signRequestId);
    if (!cached?.length) return undefined;
    const byEmail = cached.find(
      (c) => normalizeEmail(c.email) && normalizeEmail(c.email) === normalizeEmail(participant.email)
    );
    if (byEmail?.role) return byEmail.role;
    if (typeof participant.order === "number") {
      const byOrder = cached.find((c) => c.order === participant.order);
      if (byOrder?.role) return byOrder.role;
    }
    const byIndex = cached[index];
    return byIndex?.role;
  };

  const refreshActiveSignRequests = useCallback(async () => {
    setActiveSignRequestsLoading(true);
    try {
      const data = await boxAction<{
        signRequests?: Array<{
          id: string;
          name?: string;
          status: string;
          signers?: Array<{ email?: string; role?: string; order?: number; status?: string }>;
          autoExpireAt?: string;
          createdAt?: string;
        }>;
      }>("list_signature_requests");
      if (data.error) return;

      const normalized = (data.signRequests ?? [])
        .map((request) => ({
          ...request,
          signers: (request.signers ?? []).map((signer, index) => ({
            ...signer,
            role: resolveRole(request.id, signer, index),
          })),
        }))
        .filter((request) => isActiveSignRequestStatus(request.status))
        .sort(
          (a, b) =>
            new Date(b.createdAt ?? 0).getTime() - new Date(a.createdAt ?? 0).getTime()
        );

      setActiveSignRequests(normalized);
    } finally {
      setActiveSignRequestsLoading(false);
    }
  }, [setActiveSignRequests, setActiveSignRequestsLoading]);

  useEffect(() => {
    void refreshActiveSignRequests();
  }, [refreshActiveSignRequests]);

  const buildCreateSignature = (
    params: Record<string, unknown>,
    participants: Array<{ email: string; role: string }>
  ) =>
    JSON.stringify({
      fileId: String(params.fileId ?? "").trim(),
      parentFolderId: String(params.parentFolderId ?? "").trim(),
      participants,
      isSequential: Boolean(params.isSequential),
      daysValid:
        params.daysValid == null || params.daysValid === ""
          ? null
          : Number(params.daysValid),
      areRemindersEnabled: Boolean(params.areRemindersEnabled),
      name: params.name != null ? String(params.name) : "",
      emailSubject: params.emailSubject != null ? String(params.emailSubject) : "",
      emailMessage: params.emailMessage != null ? String(params.emailMessage) : "",
    });

  const stagePreviewAndDetails = async (
    params: Record<string, unknown>,
    participants: Array<{ email: string; role: string }>
  ) => {
    const fileId = params.fileId ? String(params.fileId).trim() : "";
    if (!fileId) throw new Error("fileId is required.");
    const previewRes = await boxAction<{
      fileId?: string;
      fileName?: string;
      token?: string;
      embedUrl?: string;
    }>(
      "get_file_preview_info",
      { fileId }
    );
    if (previewRes.error) {
      throw new Error(chatError(previewRes, "Could not load document preview."));
    }
    const token = previewRes.token;
    const fileName = previewRes.fileName ?? "Document";
    const embedUrl = previewRes.embedUrl;
    if (!token && !embedUrl) throw new Error("Preview source is not available.");

    const requestSummary = {
      fileId,
      fileName,
      participants,
      isSequential: !!params.isSequential,
      daysValid: params.daysValid != null ? Number(params.daysValid) : undefined,
      areRemindersEnabled: !!params.areRemindersEnabled,
      name: params.name ? String(params.name) : undefined,
      emailSubject: params.emailSubject ? String(params.emailSubject) : undefined,
    };

    setPreviewData({
      fileId,
      fileName,
      token: token ?? "",
      embedUrl,
      requestSummary,
    });

    return { fileName, requestSummary };
  };

  useCopilotAction({
    name: "search_files",
    description:
      "Search for files in the user's Box account by name or description. Use this when the user asks to find a document, template, contract, NDA, or any file.",
    parameters: [
      {
        name: "query",
        type: "string",
        description: "Search query (e.g. 'NDA template', 'vendor agreement')",
        required: true,
      },
    ],
    handler: async ({ query }) => {
      const q = String(query ?? "").trim();
      if (!q) throw new Error("Search query is required.");
      const data = await boxAction<{ files?: Array<{ id: string; name?: string; parentId?: string }> }>("search_files", { q });
      if (data.error) throw new Error(chatError(data, "Search failed."));
      const files = data.files ?? [];
      if (!files.length)
        return `No files found for "${query}". Suggest trying a different search or uploading the document to Box first.`;
      return `Found ${files.length} file(s):\n${files
        .map((f) => `- **${f.name || "Unnamed"}** — fileId: \`${f.id}\`${f.parentId ? `, parentFolderId: \`${f.parentId}\`` : ""}`)
        .join("\n")}\n\nWhen creating a signature request, use the exact fileId (and optionally parentFolderId) from above.`;
    },
  });

  useCopilotAction({
    name: "prepare_signature_request",
    description:
      "Prepare a signature request by showing the user a document preview and request details. Call this FIRST before create_signature_request. Use the same parameters (fileId, participants, etc.). After the user confirms in chat (e.g. 'yes' or 'confirm'), call create_signature_request with the same parameters.",
    parameters: [
      { name: "fileId", type: "string", description: "Box file ID to be signed", required: true },
      { name: "parentFolderId", type: "string", description: "Box folder ID for signed document. Optional.", required: false },
      { name: "approverEmails", type: "string[]", description: "Emails of approvers.", required: false },
      { name: "signerEmails", type: "string[]", description: "Emails of signers.", required: false },
      { name: "finalCopyReaderEmails", type: "string[]", description: "Emails for final copy.", required: false },
      { name: "participants", type: "object[]", description: "Array of { email, role }. Optional.", required: false },
      { name: "isSequential", type: "boolean", description: "Sequential signing order.", required: false },
      { name: "daysValid", type: "number", description: "Days until expiration.", required: false },
      { name: "areRemindersEnabled", type: "boolean", description: "Enable reminders.", required: false },
      { name: "name", type: "string", description: "Request name.", required: false },
      { name: "emailSubject", type: "string", description: "Email subject.", required: false },
      { name: "emailMessage", type: "string", description: "Email message body.", required: false },
    ],
    handler: async (params) => {
      const participants = buildParticipantsFromParams({
        participantsRaw: params.participants,
        approverEmailsRaw: params.approverEmails,
        signerEmailsRaw: params.signerEmails,
        finalCopyReaderEmailsRaw: params.finalCopyReaderEmails,
      });
      const fileId = params.fileId ? String(params.fileId).trim() : "";
      if (!fileId) throw new Error("fileId is required.");
      if (!participants.length)
        throw new Error("Specify who is involved: participants or approverEmails/signerEmails.");
      const { fileName, requestSummary } = await stagePreviewAndDetails(
        params as Record<string, unknown>,
        participants
      );
      pendingCreateSignatureRef.current = buildCreateSignature(
        params as Record<string, unknown>,
        participants
      );
      return `I've prepared the signature request. The document **${fileName}** is shown in the preview panel next to this chat, with the request details below it. Please review and confirm: **Is this the correct document to sign?** Reply **Yes** or **Confirm** to send the signature request; I will then create it with the same parameters (${participants.length} participant(s), ${requestSummary.isSequential ? "sequential signing" : "any order"}${requestSummary.daysValid ? `, expires in ${requestSummary.daysValid} days` : ""}${requestSummary.areRemindersEnabled ? ", reminders enabled" : ""}).`;
    },
  });

  useCopilotAction({
    name: "create_signature_request",
    description:
      "Create a Box Sign signature request. Required: fileId and at least one participant. For workflows with BOTH an approver and a signer: pass approverEmails (array of emails) and signerEmails (array of emails) so roles are correct; set isSequential true. Order is always: approvers first, then signers, then final copy readers. For signers only use signerEmails. Use participants only when you need a custom order with explicit role per person.",
    parameters: [
      {
        name: "fileId",
        type: "string",
        description: "Box file ID to be signed (get from search_files)",
        required: true,
      },
      {
        name: "parentFolderId",
        type: "string",
        description:
          "Box folder ID where the signed document will be stored. Get from search_files result (parentId). Optional: if omitted, the file's parent folder is used.",
        required: false,
      },
      {
        name: "approverEmails",
        type: "string[]",
        description: "Emails of approvers (review/approve only, do not sign). Use when user says 'approver' or 'for approval'. For 'one approver then one signer' pass this AND signerEmails.",
        required: false,
      },
      {
        name: "signerEmails",
        type: "string[]",
        description: "Emails of signers (must sign the document). Use when user says 'signer' or 'pass it to [X] to sign'. For 'approver then signer' pass this AND approverEmails.",
        required: false,
      },
      {
        name: "finalCopyReaderEmails",
        type: "string[]",
        description: "Emails of people who only receive a copy when done (no action). Optional.",
        required: false,
      },
      {
        name: "participants",
        type: "object[]",
        description:
          "Alternative: array of { email: string, role: 'signer'|'approver'|'final_copy_reader' }. Use only when not using approverEmails/signerEmails. Order = sequence when isSequential.",
        required: false,
      },
      {
        name: "isSequential",
        type: "boolean",
        description: "If true, participants act in order (first in list first). Default false.",
        required: false,
      },
      {
        name: "daysValid",
        type: "number",
        description: "Days until the request expires (0-730). Optional.",
        required: false,
      },
      {
        name: "areRemindersEnabled",
        type: "boolean",
        description: "Enable reminder cadence (days 3, 8, 13, 18); only days before expiration apply.",
        required: false,
      },
      {
        name: "name",
        type: "string",
        description: "Name of the signature request. Optional.",
        required: false,
      },
      {
        name: "emailSubject",
        type: "string",
        description: "Custom subject line for the sign request email. Optional.",
        required: false,
      },
      {
        name: "emailMessage",
        type: "string",
        description: "Custom message body for the sign request email (supports basic HTML). Optional.",
        required: false,
      },
    ],
    handler: async (params) => {
      const participants = buildParticipantsFromParams({
        participantsRaw: params.participants,
        approverEmailsRaw: params.approverEmails,
        signerEmailsRaw: params.signerEmails,
        finalCopyReaderEmailsRaw: params.finalCopyReaderEmails,
      });
      const {
        fileId,
        parentFolderId,
        isSequential,
        daysValid,
        areRemindersEnabled,
        name,
        emailSubject,
        emailMessage,
      } = params;
      const approverEmails = Array.isArray(params.approverEmails) ? params.approverEmails.map((e) => String(e).trim()).filter(Boolean) : params.approverEmails != null ? [String(params.approverEmails).trim()].filter(Boolean) : [];
      const signerEmails = Array.isArray(params.signerEmails) ? params.signerEmails.map((e) => String(e).trim()).filter(Boolean) : params.signerEmails != null ? [String(params.signerEmails).trim()].filter(Boolean) : [];
      const finalCopyReaderEmails = Array.isArray(params.finalCopyReaderEmails) ? params.finalCopyReaderEmails.map((e) => String(e).trim()).filter(Boolean) : params.finalCopyReaderEmails != null ? [String(params.finalCopyReaderEmails).trim()].filter(Boolean) : [];
      const hasRoleLists = approverEmails.length > 0 || signerEmails.length > 0 || finalCopyReaderEmails.length > 0;
      if (!fileId)
        throw new Error("fileId is required.");
      if (!participants.length)
        throw new Error(
          "Specify who is involved: use participants (array of { email, role }), or approverEmails and signerEmails, e.g. approverEmails: ['legal@company.com'], signerEmails: ['manager@company.com']."
        );

      const currentSignature = buildCreateSignature(
        params as Record<string, unknown>,
        participants
      );
      if (pendingCreateSignatureRef.current !== currentSignature) {
        const { fileName, requestSummary } = await stagePreviewAndDetails(
          params as Record<string, unknown>,
          participants
        );
        pendingCreateSignatureRef.current = currentSignature;
        return `Before sending, I must show the document preview and signing details. I have displayed **${fileName}** with all request details on screen. Please review and confirm in chat ("Yes" or "Confirm"), then ask me to send this exact request. Details: ${participants.length} participant(s), ${requestSummary.isSequential ? "sequential signing" : "any order"}${requestSummary.daysValid ? `, expires in ${requestSummary.daysValid} days` : ""}${requestSummary.areRemindersEnabled ? ", reminders enabled" : ""}.`;
      }

      const payload: Record<string, unknown> = {
        fileId: String(fileId),
        parentFolderId: parentFolderId != null && String(parentFolderId).trim() ? String(parentFolderId).trim() : undefined,
        participants,
        isSequential: !!isSequential,
        daysValid,
        areRemindersEnabled: !!areRemindersEnabled,
        name: name ? String(name) : undefined,
        emailSubject: emailSubject ? String(emailSubject) : undefined,
        emailMessage: emailMessage ? String(emailMessage) : undefined,
      };
      if (hasRoleLists) {
        payload.approverEmails = approverEmails;
        payload.signerEmails = signerEmails;
        payload.finalCopyReaderEmails = finalCopyReaderEmails;
      }
      const data = await boxAction<{
        signRequest?: {
          id?: string;
          status?: string;
          signers?: Array<{ email?: string; role?: string }>;
          prepareUrl?: string;
        };
      }>("create_signature_request", payload);
      if (data.error) {
        const message = chatError(data, "Create failed.");
        throw new Error(message);
      }
      const sr = data.signRequest;
      if (sr?.id) {
        lastCreatedSignRequestIdRef.current = sr.id;
        signRequestRoleCacheRef.current.set(
          sr.id,
          participants.map((p, index) => ({ email: p.email, role: p.role, order: !!isSequential ? index : undefined }))
        );
      }
      pendingCreateSignatureRef.current = null;
      setPreviewData(null);
      await refreshActiveSignRequests();
      const participantLines =
        sr?.signers
          ?.map((s, index) => {
            const resolvedRole = resolveRole(sr?.id, s, index);
            return `  - ${s.email} — **${roleLabel(resolvedRole)}**`;
          })
          .join("\n") ?? "";
      return `Signature request created successfully.\n- **ID**: ${sr?.id}\n- **Status**: ${sr?.status}\n- **Participants (by role):**\n${participantLines || "  —"}\n${sr?.prepareUrl ? `- **Prepare URL** (add signature fields): ${sr.prepareUrl}` : ""}`;
    },
  });

  useCopilotAction({
    name: "bulk_prepare_signature_requests",
    description:
      "Prepare multiple signature requests by showing document previews that the user can browse through. Call this FIRST and ONLY ONCE when user requests multiple sign requests. After showing previews, WAIT for user confirmation before calling bulk_create_signature_requests.",
    parameters: [
      {
        name: "requests",
        type: "object[]",
        description: "Array of signature request configurations. Each request should have: fileId (required), participants or approverEmails/signerEmails (required), and optional: parentFolderId, isSequential, daysValid, areRemindersEnabled, name, emailSubject, emailMessage",
        required: true,
      },
    ],
    handler: async ({ requests }) => {
      if (!Array.isArray(requests) || requests.length === 0) {
        throw new Error("requests array is required and must contain at least one request.");
      }

      const previewDocuments: Array<{
        fileId: string;
        fileName: string;
        token: string;
        embedUrl?: string;
        requestSummary: RequestSummary;
      }> = [];

      for (const req of requests) {
        try {
          const participants = buildParticipantsFromParams({
            participantsRaw: req.participants,
            approverEmailsRaw: req.approverEmails,
            signerEmailsRaw: req.signerEmails,
            finalCopyReaderEmailsRaw: req.finalCopyReaderEmails,
          });

          const fileId = req.fileId ? String(req.fileId).trim() : "";
          if (!fileId) {
            throw new Error("fileId is required for each request.");
          }

          if (!participants.length) {
            throw new Error(`Request for file ${fileId} requires at least one participant.`);
          }

          const previewRes = await boxAction<{
            fileId?: string;
            fileName?: string;
            token?: string;
            embedUrl?: string;
          }>("get_file_preview_info", { fileId });

          if (previewRes.error) {
            throw new Error(chatError(previewRes, `Could not load preview for file ${fileId}.`));
          }

          const token = previewRes.token;
          const fileName = previewRes.fileName ?? "Document";
          const embedUrl = previewRes.embedUrl;

          if (!token && !embedUrl) {
            throw new Error(`Preview source is not available for file ${fileId}.`);
          }

          previewDocuments.push({
            fileId,
            fileName,
            token: token ?? "",
            embedUrl,
            requestSummary: {
              fileId,
              fileName,
              participants,
              isSequential: !!req.isSequential,
              daysValid: req.daysValid != null ? Number(req.daysValid) : undefined,
              areRemindersEnabled: !!req.areRemindersEnabled,
              name: req.name ? String(req.name) : undefined,
              emailSubject: req.emailSubject ? String(req.emailSubject) : undefined,
            },
          });
        } catch (error) {
          throw new Error(`Failed to prepare preview: ${error instanceof Error ? error.message : String(error)}`);
        }
      }

      // Store the pending requests for later creation
      pendingCreateSignatureRef.current = JSON.stringify(requests);

      // Set preview data with multiple documents
      setPreviewData({ documents: previewDocuments });

      const fileList = previewDocuments.map((doc, idx) => `${idx + 1}. **${doc.fileName}**`).join("\n");
      
      return `I've prepared ${previewDocuments.length} signature request${previewDocuments.length > 1 ? "s" : ""}. You can browse through all documents using the navigation buttons in the preview panel:\n\n${fileList}\n\nPlease review all documents and their details. Reply **Yes** or **Confirm** in the chat to send all ${previewDocuments.length} signature requests.`;
    },
  });

  useCopilotAction({
    name: "bulk_create_signature_requests",
    description:
      "Create multiple Box Sign signature requests at once. ONLY call this AFTER calling bulk_prepare_signature_requests AND receiving user confirmation. This actually creates all the signature requests in Box. Call this exactly ONCE per bulk operation.",
    parameters: [
      {
        name: "requests",
        type: "object[]",
        description: "Array of signature request configurations. Each request should have: fileId (required), participants or approverEmails/signerEmails (required), and optional: parentFolderId, isSequential, daysValid, areRemindersEnabled, name, emailSubject, emailMessage",
        required: true,
      },
    ],
    handler: async ({ requests }) => {
      if (!Array.isArray(requests) || requests.length === 0) {
        throw new Error("requests array is required and must contain at least one request.");
      }

      // Check if bulk_prepare was called first
      const pendingRequests = pendingCreateSignatureRef.current;
      const hasPending = pendingRequests !== null;
      
      // Clear preview state to avoid conflicts
      pendingCreateSignatureRef.current = null;
      setPreviewData(null);

      // If no pending requests, this might be called without prepare (which is allowed, but note it)
      if (!hasPending) {
        console.log("[bulk_create] Called without prior bulk_prepare - proceeding directly");
      }

      const results: Array<{
        success: boolean;
        fileId: string;
        requestId?: string;
        participants?: Array<{ email: string; role: string }>;
        error?: string;
      }> = [];

      for (const req of requests) {
        try {
          const participants = buildParticipantsFromParams({
            participantsRaw: req.participants,
            approverEmailsRaw: req.approverEmails,
            signerEmailsRaw: req.signerEmails,
            finalCopyReaderEmailsRaw: req.finalCopyReaderEmails,
          });

          const fileId = req.fileId ? String(req.fileId).trim() : "";
          if (!fileId) {
            results.push({
              success: false,
              fileId: fileId || "unknown",
              error: "fileId is required",
            });
            continue;
          }

          if (!participants.length) {
            results.push({
              success: false,
              fileId,
              error: "At least one participant is required",
            });
            continue;
          }

          const approverEmails = Array.isArray(req.approverEmails) ? req.approverEmails.map((e: string) => String(e).trim()).filter(Boolean) : req.approverEmails != null ? [String(req.approverEmails).trim()].filter(Boolean) : [];
          const signerEmails = Array.isArray(req.signerEmails) ? req.signerEmails.map((e: string) => String(e).trim()).filter(Boolean) : req.signerEmails != null ? [String(req.signerEmails).trim()].filter(Boolean) : [];
          const finalCopyReaderEmails = Array.isArray(req.finalCopyReaderEmails) ? req.finalCopyReaderEmails.map((e: string) => String(e).trim()).filter(Boolean) : req.finalCopyReaderEmails != null ? [String(req.finalCopyReaderEmails).trim()].filter(Boolean) : [];
          const hasRoleLists = approverEmails.length > 0 || signerEmails.length > 0 || finalCopyReaderEmails.length > 0;

          const payload: Record<string, unknown> = {
            fileId,
            parentFolderId: req.parentFolderId != null && String(req.parentFolderId).trim() ? String(req.parentFolderId).trim() : undefined,
            participants,
            isSequential: !!req.isSequential,
            daysValid: req.daysValid,
            areRemindersEnabled: !!req.areRemindersEnabled,
            name: req.name ? String(req.name) : undefined,
            emailSubject: req.emailSubject ? String(req.emailSubject) : undefined,
            emailMessage: req.emailMessage ? String(req.emailMessage) : undefined,
          };

          if (hasRoleLists) {
            payload.approverEmails = approverEmails;
            payload.signerEmails = signerEmails;
            payload.finalCopyReaderEmails = finalCopyReaderEmails;
          }

          const data = await boxAction<{
            signRequest?: {
              id?: string;
              status?: string;
              signers?: Array<{ email?: string; role?: string }>;
              prepareUrl?: string;
            };
          }>("create_signature_request", payload);

          if (data.error) {
            const errorMessage = chatError(data, "Create failed");
            console.error(`[bulk_create] Failed to create request for file ${fileId}:`, errorMessage, data);
            results.push({
              success: false,
              fileId,
              participants,
              error: errorMessage,
            });
            continue;
          }

          const sr = data.signRequest;
          if (sr?.id) {
            lastCreatedSignRequestIdRef.current = sr.id;
            signRequestRoleCacheRef.current.set(
              sr.id,
              participants.map((p, index) => ({ email: p.email, role: p.role, order: !!req.isSequential ? index : undefined }))
            );
          }

          results.push({
            success: true,
            fileId,
            requestId: sr?.id,
            participants,
          });
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : String(error);
          console.error(`[bulk_create] Exception creating request for file ${req.fileId}:`, error);
          results.push({
            success: false,
            fileId: req.fileId || "unknown",
            error: errorMessage,
          });
        }
      }

      await refreshActiveSignRequests();

      const successful = results.filter((r) => r.success);
      const failed = results.filter((r) => !r.success);

      console.log(`[bulk_create] Results: ${successful.length} succeeded, ${failed.length} failed out of ${requests.length} total`);

      let summary = `✓ Bulk signature request creation completed.\n\n`;
      summary += `**Summary**: ${successful.length} of ${requests.length} signature request${requests.length > 1 ? "s" : ""} created successfully.\n`;
      
      if (successful.length > 0) {
        summary += `\n**Successfully created:**\n`;
        successful.forEach((r, idx) => {
          const participantSummary = r.participants?.map(p => `${p.email} (${roleLabel(p.role)})`).join(", ") || "unknown participants";
          summary += `${idx + 1}. File ID: ${r.fileId}\n   - Request ID: ${r.requestId}\n   - Participants: ${participantSummary}\n`;
        });
      }

      if (failed.length > 0) {
        summary += `\n**Failed to create:**\n`;
        failed.forEach((r, idx) => {
          const participantDetails = r.participants?.map(p => `${p.email} (${roleLabel(p.role)})`).join(", ") || "no participants";
          summary += `${idx + 1}. File ID: ${r.fileId}\n   - Intended for: ${participantDetails}\n   - Error: ${r.error}\n`;
        });
        
        summary += `\n**Common causes of failures:**\n`;
        summary += `- File is in the root folder (Box Sign requires files to be in a subfolder)\n`;
        summary += `- File ID is incorrect or file doesn't exist\n`;
        summary += `- Invalid email addresses\n`;
        summary += `- Insufficient permissions\n`;
      }

      summary += `\n✓ All signature requests have been processed. The operation is complete.`;

      return summary;
    },
  });

  useCopilotAction({
    name: "list_signature_requests",
    description:
      "List the user's Box Sign signature requests. Use when the user asks for status, 'my requests', 'what's pending', etc.",
    parameters: [],
    handler: async () => {
      const data = await boxAction<{
        signRequests?: Array<{
          id: string;
          name?: string;
          status: string;
          signers?: Array<{ email?: string; role?: string; status?: string }>;
          autoExpireAt?: string;
          createdAt?: string;
        }>;
      }>("list_signature_requests");
      if (data.error) throw new Error(chatError(data, "List failed."));
      const list = data.signRequests ?? [];
      const activeList = list.filter((request) =>
        isActiveSignRequestStatus(request.status)
      );
      setActiveSignRequests(
        activeList
          .map((request) => ({
            ...request,
            signers: request.signers ?? [],
          }))
          .sort(
            (a, b) =>
              new Date(b.createdAt ?? 0).getTime() - new Date(a.createdAt ?? 0).getTime()
          )
      );
      if (!list.length) return "You have no signature requests.";
      const formatParticipant = (
        signRequestId: string,
        s: { email?: string; role?: string; order?: number; status?: string },
        index: number
      ) => {
        const resolvedRole = resolveRole(signRequestId, s, index);
        return `${s.email} — **${roleLabel(resolvedRole)}** — ${s.status || "pending"}`;
      };
      return `You have ${list.length} signature request(s):\n${list
        .map(
          (r) =>
            `- **${r.name || r.id}** (ID: \`${r.id}\`)\n  Status: ${r.status}\n  Participants: ${(r.signers || []).map((s, index) => formatParticipant(r.id, s, index)).join("; ")}${r.autoExpireAt ? `\n  Expires: ${r.autoExpireAt}` : ""}`
        )
        .join("\n\n")}`;
    },
  });

  useCopilotAction({
    name: "get_signature_request_status",
    description:
      "Get the status/details of a specific Box Sign request by ID. If ID is not provided, this action resolves the latest request in this chat session.",
    parameters: [
      {
        name: "signRequestId",
        type: "string",
        description: "The Box Sign request ID. Optional: if omitted, use the latest created request in this chat session.",
        required: false,
      },
    ],
    handler: async ({ signRequestId }) => {
      let resolvedSignRequestId = signRequestId ? String(signRequestId) : lastCreatedSignRequestIdRef.current;
      if (!resolvedSignRequestId) {
        const listData = await boxAction<{
          signRequests?: Array<{ id: string }>;
        }>("list_signature_requests");
        if (listData.error) throw new Error(chatError(listData, "List failed."));
        resolvedSignRequestId = listData.signRequests?.[0]?.id ?? null;
      }
      if (!resolvedSignRequestId) {
        throw new Error("No signature requests found. Create one first or provide signRequestId.");
      }
      const r = await boxAction<{
        id?: string;
        name?: string;
        status?: string;
        signers?: Array<{ email?: string; role?: string; order?: number; status?: string }>;
        autoExpireAt?: string;
        prepareUrl?: string;
      }>("get_signature_request_status", { signRequestId: resolvedSignRequestId });
      if (r.error) throw new Error(chatError(r, "Get failed."));
      if (r.id) {
        lastCreatedSignRequestIdRef.current = r.id;
      }
      const parts = (r.signers || []).map((s, index) => {
        const resolvedRole = resolveRole(r.id, s, index);
        return `  - ${s.email} — **${roleLabel(resolvedRole)}** — ${s.status || "pending"}`;
      });
      return `**${r.name || r.id}**\n- Status: ${r.status}\n- **Participants (role and status):**\n${parts.join("\n") || "  —"}${r.autoExpireAt ? `\n- Expires: ${r.autoExpireAt}` : ""}${r.prepareUrl ? `\n- Prepare URL: ${r.prepareUrl}` : ""}`;
    },
  });

  useCopilotAction({
    name: "get_latest_signature_request_status",
    description:
      "Get details for the latest sign request created in this chat session. Use for prompts like 'return details of this sign request'.",
    parameters: [],
    handler: async () => {
      let resolvedSignRequestId = lastCreatedSignRequestIdRef.current;
      if (!resolvedSignRequestId) {
        const listData = await boxAction<{
          signRequests?: Array<{ id: string }>;
        }>("list_signature_requests");
        if (listData.error) throw new Error(chatError(listData, "List failed."));
        resolvedSignRequestId = listData.signRequests?.[0]?.id ?? null;
      }
      if (!resolvedSignRequestId) {
        throw new Error("No signature requests found. Create one first.");
      }
      const r = await boxAction<{
        id?: string;
        name?: string;
        status?: string;
        signers?: Array<{ email?: string; role?: string; order?: number; status?: string }>;
        autoExpireAt?: string;
        prepareUrl?: string;
      }>("get_signature_request_status", { signRequestId: resolvedSignRequestId });
      if (r.error) throw new Error(chatError(r, "Get failed."));
      if (r.id) {
        lastCreatedSignRequestIdRef.current = r.id;
      }
      const parts = (r.signers || []).map((s, index) => {
        const resolvedRole = resolveRole(r.id, s, index);
        return `  - ${s.email} — **${roleLabel(resolvedRole)}** — ${s.status || "pending"}`;
      });
      return `**${r.name || r.id}**\n- Status: ${r.status}\n- **Participants (role and status):**\n${parts.join("\n") || "  —"}${r.autoExpireAt ? `\n- Expires: ${r.autoExpireAt}` : ""}${r.prepareUrl ? `\n- Prepare URL: ${r.prepareUrl}` : ""}`;
    },
  });

  useCopilotAction({
    name: "bulk_cancel_signature_requests",
    description:
      "Cancel all active sign requests that match a phrase in request name or ID. Use for commands like 'remove all sign requests for Vendor agreement'.",
    parameters: [
      {
        name: "query",
        type: "string",
        description:
          "Case-insensitive search phrase used to match request name or request ID.",
        required: true,
      },
    ],
    handler: async ({ query }) => {
      const q = String(query ?? "").trim().toLowerCase();
      if (!q) throw new Error("query is required.");

      const listData = await boxAction<{
        signRequests?: Array<{
          id: string;
          name?: string;
          status: string;
        }>;
      }>("list_signature_requests");
      if (listData.error) throw new Error(chatError(listData, "List failed."));

      const all = listData.signRequests ?? [];
      const matches = all.filter((request) => {
        const name = (request.name ?? "").toLowerCase();
        const id = request.id.toLowerCase();
        return name.includes(q) || id.includes(q);
      });

      if (!matches.length) {
        return `No signature requests matched "${query}".`;
      }

      const activeMatches = matches.filter((request) =>
        isActiveSignRequestStatus(request.status)
      );
      const skipped = matches.filter((request) => !isActiveSignRequestStatus(request.status));

      let cancelled = 0;
      const cancelledIds: string[] = [];
      const failed: Array<{ id: string; error: string }> = [];
      for (const request of activeMatches) {
        const cancelData = await boxAction<{ error?: string }>("cancel_signature_request", {
          signRequestId: request.id,
        });
        if (cancelData.error) {
          failed.push({ id: request.id, error: cancelData.error });
        } else {
          cancelled += 1;
          cancelledIds.push(request.id);
        }
      }

      await refreshActiveSignRequests();

      const cancelledLines = activeMatches
        .filter((request) => cancelledIds.includes(request.id))
        .map((request) => `- ${request.name || request.id} (\`${request.id}\`)`);

      const failedLines = failed.map((f) => `- \`${f.id}\`: ${f.error}`);
      const skippedLines = skipped.map(
        (request) => `- ${request.name || request.id} (\`${request.id}\`) — ${request.status}`
      );

      return [
        `Bulk cancellation result for "${query}":`,
        `- Matched: ${matches.length}`,
        `- Cancelled: ${cancelled}`,
        `- Skipped (already inactive): ${skipped.length}`,
        `- Failed: ${failed.length}`,
        cancelledLines.length ? `\nCancelled:\n${cancelledLines.join("\n")}` : "",
        skippedLines.length ? `\nSkipped:\n${skippedLines.join("\n")}` : "",
        failedLines.length ? `\nFailed:\n${failedLines.join("\n")}` : "",
      ]
        .filter(Boolean)
        .join("\n");
    },
  });

  useCopilotAction({
    name: "cancel_signature_request",
    description:
      "Cancel a pending Box Sign request. Use when the user wants to cancel a request by ID.",
    parameters: [
      {
        name: "signRequestId",
        type: "string",
        description: "The Box Sign request ID to cancel",
        required: true,
      },
    ],
    handler: async ({ signRequestId }) => {
      if (!signRequestId) throw new Error("signRequestId is required.");
      const data = await boxAction("cancel_signature_request", { signRequestId: String(signRequestId) });
      if (data.error) throw new Error(chatError(data, "Cancel failed."));
      await refreshActiveSignRequests();
      return `Signature request \`${signRequestId}\` has been cancelled.`;
    },
  });

  useCopilotAction({
    name: "resend_signature_request",
    description:
      "Resend the signature request email to all outstanding signers. Use when the user wants to resend reminders.",
    parameters: [
      {
        name: "signRequestId",
        type: "string",
        description: "The Box Sign request ID to resend",
        required: true,
      },
    ],
    handler: async ({ signRequestId }) => {
      if (!signRequestId) throw new Error("signRequestId is required.");
      const data = await boxAction("resend_signature_request", { signRequestId: String(signRequestId) });
      if (data.error) throw new Error(chatError(data, "Resend failed."));
      return `Signature request \`${signRequestId}\` emails have been resent to outstanding signers.`;
    },
  });

  return null;
}
