"use client";

import { CopilotKit, useCopilotAction, useCopilotAdditionalInstructions } from "@copilotkit/react-core";
import { CopilotSidebar } from "@copilotkit/react-ui";
import "@copilotkit/react-ui/styles.css";
import { FileSignature, Search, ListChecks } from "lucide-react";
import { createContext, useContext, useRef, useState, useEffect, useCallback } from "react";
import { BoxPreviewPanel, type PreviewData, type RequestSummary } from "./BoxPreviewPanel";

const AGENT_INSTRUCTIONS = `You are the Box Sign AI Assistant. You help users create and manage e-signature requests in Box.

🚨🚨🚨 READ THIS FIRST - CRITICAL WORKFLOW 🚨🚨🚨

Every signature request creation follows these 4 STEPS IN ORDER:
1. Extract file ID + participants
2. Call confirm_security_preferences action ← YOU MUST DO THIS
3. Call prepare_signature_request (after user responds)
4. Call create_signature_request (after user confirms)

YOU CANNOT SKIP STEP 2. If you call prepare_signature_request without first calling confirm_security_preferences, you are violating the workflow.

⚠️⚠️⚠️ MANDATORY 4-STEP WORKFLOW - NO SKIPPING ALLOWED ⚠️⚠️⚠️

When a user requests to create a signature request, you MUST follow these 4 steps IN ORDER:

**STEP 1: Extract and Validate**
- Extract file ID and participants from user's request
- If missing, ask for them
- Once you have both, GO TO STEP 2

**STEP 2: Call confirm_security_preferences Action (MANDATORY - CANNOT SKIP)**
- ⚠️ THIS IS REQUIRED - YOU CANNOT PROCEED TO STEP 3 WITHOUT DOING THIS ⚠️
- Call the confirm_security_preferences action with:
  - fileId: the file ID you have
  - participantEmails: array of all participant emails
  - requestSummary: brief summary of options (if any)
- This action returns a message asking the user about security features
- Display the returned message to the user
- **STOP HERE** - Do NOT call prepare_signature_request
- WAIT for the user's response in their next message

**STEP 3: Call prepare_signature_request (ONLY After user responds to security question)**
- User has now answered the security question
- Call prepare_signature_request with all parameters
- Include securityDecision with one of: phone_verification, password_protection, box_login_required, multiple, or none
- Include any security features they requested
- Tell user to review and confirm the preview

**STEP 4: Call create_signature_request (After user confirms preview)**
- User says "yes" or "confirm"
- Call create_signature_request with the same parameters (including securityDecision)

⚠️ CRITICAL RULE ⚠️
IF YOU HAVE FILE ID + PARTICIPANTS → YOU MUST CALL confirm_security_preferences BEFORE prepare_signature_request
DO NOT SKIP STEP 2 UNDER ANY CIRCUMSTANCES
IF YOU CALL prepare_signature_request WITHOUT FIRST CALLING confirm_security_preferences, YOU ARE DOING IT WRONG

## Scope: document signing only
- You must **only** answer questions and perform actions related to **document signing with Box Sign**: searching files in Box, creating signature requests, listing or checking status of sign requests, cancelling or resending requests, and explaining how Box Sign options work (expiration, reminders, signer order, email subject/message).
- For **any other topic** (weather, news, general knowledge, other products, off-topic chat), do **not** answer or use tools. Reply briefly and redirect: "I can only help with Box Sign—searching documents, creating and managing signature requests. Is there something you'd like to do with a document to sign?" Do not provide information about weather, locations, or anything unrelated to the signing process.

## CRITICAL: Preview and confirm before sending (STEP 3 and STEP 4)
- **STEP 3**: Call **prepare_signature_request** (but only AFTER you called confirm_security_preferences in STEP 2). Include securityDecision based on the user's reply (phone_verification, password_protection, box_login_required, multiple, or none). This shows the user a document preview and request details next to the chat. Tell the user: "I've prepared the request. You can see the document preview and details. Is this the correct document to sign? Reply **Yes** or **Confirm** to send the signature request."
- **STEP 4**: **Only after** the user explicitly confirms (e.g. "yes", "confirm", "looks good", "send it", "go ahead") call **create_signature_request** with the same parameters you used in prepare_signature_request (including securityDecision). Never call create_signature_request without having first called prepare_signature_request and received user confirmation.

## Handling user-requested changes to the preview
- **If the user requests changes** after seeing the preview (e.g., "change expiration to 14 days", "add another signer", "make it sequential", "change email subject"), **immediately call prepare_signature_request again** with the updated parameters.
- The UI will automatically update to show the new request details.
- After calling prepare_signature_request with the updated parameters, tell the user: "I've updated the request details. The preview panel now shows [describe the changes]. Please review and confirm if this looks correct."
- **Important**: Extract and preserve all previous parameters, only modifying what the user requested to change. For example, if they had 2 signers and 7-day expiration, and they ask to "enable reminders", keep the signers and expiration, just add areRemindersEnabled: true.
- You can call prepare_signature_request multiple times as the user refines their request. Each call updates the preview panel.
- Only call create_signature_request after the user explicitly confirms the final version.

## Handling combined / single-message requests
Users often give many details in one message. **Parse and extract** everything they said, then follow the 4-STEP WORKFLOW described at the top.

⚠️ CRITICAL REMINDER ⚠️
After you extract file ID + participants, DO NOT call prepare_signature_request yet.
Your NEXT ACTION must be: confirm_security_preferences
This action asks the user about security and returns a message.
WAIT for user's response, THEN call prepare_signature_request.

**Example 1 - Simple request:** "One signer, alice@company.com, valid for a week, please enable reminder one day before expiration, also add subject '[NDA] Important document from Box'"

User provides: participants, daysValid, areRemindersEnabled, emailSubject
Missing: file ID

**Agent Response 1:** "Which document should I use for this signature request? You can give me a file ID or I can search Box by name."

User: "file 123456"

**Agent Response 2 - Call confirm_security_preferences:**
Call confirm_security_preferences with:
- fileId: "123456"
- participantEmails: ["alice@company.com"]
- requestSummary: "7 days expiration, reminders enabled, custom email subject"

Action returns security question message to user. WAIT for user's response.

User: "No, proceed"

**Agent Response 3 - Call prepare_signature_request:**
Call prepare_signature_request with: fileId: "123456", participants: [{ email: "alice@company.com", role: "signer" }], daysValid: 7, areRemindersEnabled: true, emailSubject: "[NDA] Important document from Box"

Then say: "Please review and confirm."

**Example 2 - Advanced request showing WORKFLOW:**

User: "Create a sign request for file 123456 with john@company.com and jane@company.com, sequential, expire in 30 days, subject 'Contract to sign'"

**Agent STEP 1 - Extract and validate:**
- fileId: "123456" ✓
- participants: john@company.com, jane@company.com ✓  
- options: isSequential: true, daysValid: 30, emailSubject: "Contract to sign" ✓
- Proceed to STEP 2

**Agent STEP 2 - Call confirm_security_preferences:**
Call confirm_security_preferences action with:
- fileId: "123456"
- participantEmails: ["john@company.com", "jane@company.com"]
- requestSummary: "sequential signing, 30 days expiration, custom email subject"

Action returns message to user asking about security. Agent displays this and STOPS.

User: "No thanks, just proceed"

**Agent STEP 3 - Call prepare_signature_request:**
Call prepare_signature_request with: fileId: "123456", participants: [{ email: "john@company.com", role: "signer" }, { email: "jane@company.com", role: "signer" }], isSequential: true, daysValid: 30, emailSubject: "Contract to sign"

Tell user: "Please review the document and details in the preview panel."

User: "Yes, looks good"

**Agent STEP 4 - Call create_signature_request:**
Call create_signature_request with same parameters

KEY TAKEAWAY FROM EXAMPLES: Always call confirm_security_preferences action after you have file ID and participants. This action returns the security question to the user. Never skip this step.

**Example 3 - With approver:** "Send the NDA for signature to alice@co.com, and legal@co.com should approve it; expire in 14 days"
- Extract: participants = [{ email: "alice@co.com", role: "signer" }, { email: "legal@co.com", role: "approver" }], daysValid = 14. Do **not** put legal@co.com as a signer—they are an approver.
- Missing: file ID → ask for it
- Once you have file ID, call confirm_security_preferences with the file ID and both participant emails
- Wait for user's security response, then call prepare_signature_request

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
- **Phone verification:** "verify by phone for [email]" / "phone verification +1234567890 for alice@co.com" → Add verificationPhoneNumber: "+1234567890" to that participant in participants array. Format must include country code (e.g., "+12125551234"). You MUST use participants array (not approverEmails/signerEmails) when security features are specified.
- **Password protection:** "require password" / "password 'secret123' for [email]" → Add password: "secret123" to that participant in participants array. The participant will need this password to access the signing request.
- **Login required:** "require Box login for [email]" / "must be logged in" → Add loginRequired: true to that participant in participants array. The participant must be logged into their Box account to access the request.
- **Security features are per-participant.** When user requests security, use participants array with objects like: { email: "alice@co.com", role: "signer", verificationPhoneNumber: "+12125551234", password: "secret", loginRequired: true }.
- **When suggesting security features proactively**: If user hasn't mentioned security, ask: "For added security, I can require phone verification (participant enters a code sent to their phone), password protection (you set a password they need to access the request), or require Box login. Would you like any of these for your participants?"
- At least one participant is required. Prefer **approverEmails/signerEmails/finalCopyReaderEmails** for simple role-based workflows. Use **participants** array when: (1) custom ordering needed, OR (2) any security features (phone/password/login) specified.

## IMPORTANT: Email customization limitations
- **Box Sign sends ONE email subject and ONE email message to ALL participants.** There is no per-participant email customization.
- If user says "send email subject X to the approver" or "send email Y to the signer", use the LAST mentioned email as the unified subject/message, and explain: "Box Sign will send this email subject/message to all participants: [list all participant emails]. Box Sign doesn't support different emails for different participants."
- If user requests different subjects/messages for different roles, politely clarify this limitation and ask which one they'd like to use for everyone.
- The preview panel will show "All X participant(s) will receive this email subject/message" to make this clear.

## Summary of workflow

1. **User provides request** → Extract details, check for missing file/participants
2. **Call confirm_security_preferences action** → This asks user about security and returns, WAIT for user response
3. **User answers security question** → Call prepare_signature_request with all parameters (including any security they requested)
4. **User confirms preview** → Call create_signature_request

CRITICAL: You MUST call confirm_security_preferences before prepare_signature_request. Never skip this step.

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

## Understanding the active requests list
- The "Active signing requests" table shows only requests with these statuses: **converting**, **sent**, **viewed**, or **downloaded** (requests that require action or are in progress).
- Box API returns ALL historical requests, but we filter to show only those awaiting participant action.
- Requests with these statuses are NOT shown (filtered out): cancelled, completed, declined, expired, error, signed (fully signed), approved (fully approved), finalizing, created (draft/unsent).
- If a user asks why they see different requests in the app vs. Box UI: Box API may return requests from all users the token has access to, or old draft requests that were never fully sent. The Box web UI typically shows only requests relevant to that specific user.
- The list auto-refreshes after creating or cancelling requests. There may be a brief delay (1-2 seconds) for Box API to update.
- If user reports seeing requests in the app that don't exist in Box UI, these might be: (1) draft requests that were created but never sent, (2) requests from a different Box user the token has access to, or (3) requests in an error state. Suggest they check the browser console logs for request details.

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
  
  // Prefer userMessage (most user-friendly)
  if (userMessage) {
    console.log("[chatError] Using userMessage:", userMessage);
    return userMessage;
  }
  
  // Fall back to error + hint
  const hint = data?.debug?.hint?.trim();
  const combined = hint ? `${msg} ${hint}` : msg;
  console.log("[chatError] Using error + hint:", combined);
  return combined;
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

// Box Sign request statuses that ARE considered active and should be shown
// Box API returns all requests including cancelled/completed ones, so we use a whitelist approach
// Only show requests that require user attention or are in progress
// See: https://developer.box.com/reference/resources/sign-request/
const ACTIVE_SIGNATURE_STATUSES = new Set([
  "converting",        // Document is being converted
  "sent",             // Request sent to participants (waiting for action)
  "viewed",           // Participant viewed the request (waiting for action)
  "downloaded",       // Participant downloaded the document (waiting for action)
]);

function isActiveSignRequestStatus(status?: string): boolean {
  const normalizedStatus = (status ?? "").toLowerCase().trim();
  const isActive = ACTIVE_SIGNATURE_STATUSES.has(normalizedStatus);
  
  // Log for debugging
  if (!isActive && normalizedStatus) {
    console.log(`[isActiveSignRequestStatus] Filtering out status: "${normalizedStatus}"`);
  }
  
  return isActive;
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
            padding: "var(--space-xl)",
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
            marginBottom: 0,
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
          <div
            style={{
              maxWidth: "72rem",
              margin: "0 auto",
              width: "100%",
              padding: "0 var(--space-xl)",
              boxSizing: "border-box",
            }}
          >
            <ActiveSigningRequestsPanel
              requests={activeSignRequests}
              loading={activeSignRequestsLoading}
            />
          </div>
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

type ParticipantWithSecurity = {
  email: string;
  role: string;
  verificationPhoneNumber?: string;
  password?: string;
  loginRequired?: boolean;
};

type SecurityDecision =
  | "phone_verification"
  | "password_protection"
  | "box_login_required"
  | "multiple"
  | "none";

type SecurityGateState = {
  prompted: boolean;
  acknowledged: boolean;
  decision?: SecurityDecision;
  updatedAt: number;
};

function buildSecurityGateKey(fileId: string, participants: Array<ParticipantWithSecurity>): string {
  const normalizedFileId = fileId.trim();
  const participantKey = participants
    .map((p) => p.email.trim().toLowerCase())
    .sort()
    .join("|");
  return `${normalizedFileId}::${participantKey}`;
}

function buildSecurityGateFileKey(fileId: string): string {
  return `${fileId.trim()}::*`;
}

function normalizeSecurityDecision(input: unknown): SecurityDecision | null {
  if (typeof input !== "string") return null;
  const value = input.trim().toLowerCase();
  if (!value) return null;
  if (["phone", "phone_verification", "sms", "verification_phone_number"].includes(value)) {
    return "phone_verification";
  }
  if (["password", "password_protection"].includes(value)) {
    return "password_protection";
  }
  if (["login", "box_login_required", "login_required", "box login"].includes(value)) {
    return "box_login_required";
  }
  if (["multiple", "combination", "mixed"].includes(value)) {
    return "multiple";
  }
  if (["none", "no", "no_security", "proceed_without_security", "skip"].includes(value)) {
    return "none";
  }
  return null;
}

function applySecurityDecisionFallback(
  participants: Array<ParticipantWithSecurity>,
  decision: SecurityDecision
): Array<ParticipantWithSecurity> {
  const hasExplicitSecurity = participants.some(
    (p) => Boolean(p.verificationPhoneNumber || p.password || p.loginRequired)
  );
  if (hasExplicitSecurity) return participants;

  // If the user selected login-required but the model forgot to attach
  // per-participant flags, apply it to signer(s) by default.
  if (decision === "box_login_required") {
    const hasSigner = participants.some((p) => (p.role || "").toLowerCase() === "signer");
    if (hasSigner) {
      return participants.map((p) =>
        (p.role || "").toLowerCase() === "signer" ? { ...p, loginRequired: true } : p
      );
    }
    return participants.map((p) => ({ ...p, loginRequired: true }));
  }

  return participants;
}

function getMissingSecurityDetailsMessage(
  participants: Array<ParticipantWithSecurity>,
  decision: SecurityDecision
): string | null {
  const roleName = (role?: string) => {
    const normalized = (role || "participant").toLowerCase();
    if (normalized === "signer") return "Signer";
    if (normalized === "approver") return "Approver";
    if (normalized === "final_copy_reader") return "Final copy reader";
    return "Participant";
  };
  const formatTargets = (missing: Array<ParticipantWithSecurity>) =>
    missing.map((p) => `- ${roleName(p.role)}: ${p.email}`).join("\n");
  const signerEmails = participants
    .filter((p) => (p.role || "").toLowerCase() === "signer")
    .map((p) => p.email);
  const targetEmails = signerEmails.length > 0 ? signerEmails : participants.map((p) => p.email);

  if (decision === "password_protection") {
    const missing = participants
      .filter((p) => targetEmails.includes(p.email))
      .filter((p) => !p.password);
    if (missing.length > 0) {
      return `I can add password protection. I still need the password for:\n${formatTargets(missing)}\n\nReply with one message like:\n- "${missing[0].email}: MySecurePass123"\n\nOr provide structured participants with password fields.`;
    }
  }

  if (decision === "phone_verification") {
    const missing = participants
      .filter((p) => targetEmails.includes(p.email))
      .filter((p) => !p.verificationPhoneNumber);
    if (missing.length > 0) {
      return `I can add phone verification. I still need phone number(s) with country code for:\n${formatTargets(missing)}\n\nReply with one message like:\n- "${missing[0].email}: +12125551234"\n\nOr provide structured participants with verificationPhoneNumber fields.`;
    }
  }

  if (decision === "multiple") {
    const hasAnySecurity = participants.some(
      (p) => Boolean(p.verificationPhoneNumber || p.password || p.loginRequired)
    );
    if (!hasAnySecurity) {
      return "You selected multiple security features, but I still need the exact per-participant settings. Please provide participants with one or more of: verificationPhoneNumber, password, loginRequired.";
    }
  }

  return null;
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
}): Array<ParticipantWithSecurity> {
  let rawList: Array<Record<string, unknown>> = [];
  if (Array.isArray(participantsRaw)) {
    rawList = participantsRaw as Array<Record<string, unknown>>;
  } else if (participantsRaw && typeof participantsRaw === "object") {
    const obj = participantsRaw as Record<string, unknown>;
    if (typeof (obj as { email?: string }).email === "string") {
      rawList = [participantsRaw as Record<string, unknown>];
    } else {
      rawList = Object.keys(obj)
        .filter((k) => /^\d+$/.test(k))
        .sort((a, b) => Number(a) - Number(b))
        .map((k) => obj[k] as Record<string, unknown>)
        .filter((p) => p && typeof p === "object");
    }
  } else if (typeof participantsRaw === "string") {
    try {
      const parsed = JSON.parse(participantsRaw) as unknown;
      rawList = Array.isArray(parsed) ? (parsed as Array<Record<string, unknown>>) : parsed && typeof parsed === "object" && (parsed as { email?: string }).email ? [parsed as Record<string, unknown>] : [];
    } catch {
      rawList = [];
    }
  }
  const approverEmails = Array.isArray(approverEmailsRaw) ? (approverEmailsRaw as string[]).map((e) => String(e).trim()).filter(Boolean) : approverEmailsRaw != null ? [String(approverEmailsRaw).trim()].filter(Boolean) : [];
  const signerEmails = Array.isArray(signerEmailsRaw) ? (signerEmailsRaw as string[]).map((e) => String(e).trim()).filter(Boolean) : signerEmailsRaw != null ? [String(signerEmailsRaw).trim()].filter(Boolean) : [];
  const finalCopyReaderEmails = Array.isArray(finalCopyReaderEmailsRaw) ? (finalCopyReaderEmailsRaw as string[]).map((e) => String(e).trim()).filter(Boolean) : finalCopyReaderEmailsRaw != null ? [String(finalCopyReaderEmailsRaw).trim()].filter(Boolean) : [];
  const getStr = (p: Record<string, unknown>) => (x: string): string =>
    (p[x] ?? "").toString().trim();
  const participantsFromRaw = rawList
    .map((p: Record<string, unknown>) => {
      const g = getStr(p);
      const email = (g("email") || g("Email")).trim();
      if (!email) return null;
      const role = (g("role") || g("Role") || "signer").toLowerCase();
      const participant: ParticipantWithSecurity = { 
        email, 
        role: ["signer", "approver", "final_copy_reader"].includes(role) ? role : "signer" 
      };
      
      // Extract security features if present
      const phone = g("verificationPhoneNumber") || g("verification_phone_number") || g("phone");
      if (phone) participant.verificationPhoneNumber = phone;
      
      const pwd = g("password");
      if (pwd) participant.password = pwd;
      
      const loginReq = p.loginRequired || p.login_required;
      if (loginReq === true || loginReq === "true") participant.loginRequired = true;
      
      return participant;
    })
    .filter((p): p is ParticipantWithSecurity => p != null);

  const hasRoleLists = approverEmails.length > 0 || signerEmails.length > 0 || finalCopyReaderEmails.length > 0;
  if (!hasRoleLists) {
    return participantsFromRaw;
  }

  const byEmail = new Map(
    participantsFromRaw.map((p) => [p.email.trim().toLowerCase(), p] as const)
  );
  const withSecurity = (email: string, role: string): ParticipantWithSecurity => {
    const base: ParticipantWithSecurity = { email, role };
    const match = byEmail.get(email.trim().toLowerCase());
    if (!match) return base;
    return {
      ...base,
      verificationPhoneNumber: match.verificationPhoneNumber,
      password: match.password,
      loginRequired: match.loginRequired,
    };
  };

  return [
    ...approverEmails.map((email) => withSecurity(email, "approver")),
    ...signerEmails.map((email) => withSecurity(email, "signer")),
    ...finalCopyReaderEmails.map((email) => withSecurity(email, "final_copy_reader")),
  ];
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
        maxWidth: "none",
        width: "100%",
        boxSizing: "border-box",
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
          : "Showing requests awaiting action (sent, viewed, downloaded, or converting)."}
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
  const securityGateRef = useRef<Map<string, SecurityGateState>>(new Map());
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

  const refreshActiveSignRequests = useCallback(async (delayMs: number = 0) => {
    // Add optional delay to allow Box API to propagate status changes
    if (delayMs > 0) {
      await new Promise(resolve => setTimeout(resolve, delayMs));
    }
    
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
      
      if (data.error) {
        console.error("[refreshActiveSignRequests] Error fetching sign requests:", data.error);
        return;
      }

      const allRequests = data.signRequests ?? [];
      console.log(`[refreshActiveSignRequests] Received ${allRequests.length} requests from Box API`);
      
      // Group requests by status for debugging
      const statusGroups = allRequests.reduce((acc, req) => {
        const status = req.status || 'unknown';
        if (!acc[status]) acc[status] = [];
        acc[status].push(req.id);
        return acc;
      }, {} as Record<string, string[]>);
      
      console.log('[refreshActiveSignRequests] Requests by status:', statusGroups);

      const normalized = allRequests
        .map((request) => ({
          ...request,
          signers: (request.signers ?? []).map((signer, index) => ({
            ...signer,
            role: resolveRole(request.id, signer, index),
          })),
        }))
        .filter((request) => {
          const isActive = isActiveSignRequestStatus(request.status);
          if (!isActive) {
            console.log(`  [FILTERED] ${request.id} (${request.name || 'unnamed'}): status="${request.status}"`);
          } else {
            console.log(`  [ACTIVE] ${request.id} (${request.name || 'unnamed'}): status="${request.status}"`);
          }
          return isActive;
        })
        .sort(
          (a, b) =>
            new Date(b.createdAt ?? 0).getTime() - new Date(a.createdAt ?? 0).getTime()
        );

      console.log(`[refreshActiveSignRequests] ✅ Showing ${normalized.length} active request(s) in UI`);
      console.log(`[refreshActiveSignRequests] 🚫 Filtered out ${allRequests.length - normalized.length} inactive request(s)`);
      
      // Additional debug: Show ALL requests with full details
      console.group('[refreshActiveSignRequests] 📋 ALL REQUESTS FROM BOX API:');
      allRequests.forEach((req, idx) => {
        console.log(`${idx + 1}. ID: ${req.id}`);
        console.log(`   Name: ${req.name || '(unnamed)'}`);
        console.log(`   Status: "${req.status}"`);
        console.log(`   Created: ${req.createdAt || 'unknown'}`);
        console.log(`   Signers: ${req.signers?.map(s => s.email).join(', ') || 'none'}`);
        console.log(`   ---`);
      });
      console.groupEnd();
      
      setActiveSignRequests(normalized);
    } catch (error) {
      console.error("[refreshActiveSignRequests] Exception:", error);
    } finally {
      setActiveSignRequestsLoading(false);
    }
  }, [setActiveSignRequests, setActiveSignRequestsLoading]);

  useEffect(() => {
    void refreshActiveSignRequests();
  }, [refreshActiveSignRequests]);

  const buildCreateSignature = (
    params: Record<string, unknown>,
    participants: Array<ParticipantWithSecurity>
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
    participants: Array<ParticipantWithSecurity>
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
      emailMessage: params.emailMessage ? String(params.emailMessage) : undefined,
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
      if (data.error) return chatError(data, "Search failed.");
      const files = data.files ?? [];
      if (!files.length)
        return `No files found for "${query}". Suggest trying a different search or uploading the document to Box first.`;
      return `Found ${files.length} file(s):\n${files
        .map((f) => `- **${f.name || "Unnamed"}** — fileId: \`${f.id}\`${f.parentId ? `, parentFolderId: \`${f.parentId}\`` : ""}`)
        .join("\n")}\n\nNext step: Call confirm_security_preferences with the fileId and participant emails.`;
    },
  });

  useCopilotAction({
    name: "confirm_security_preferences",
    description:
      "🚨 MANDATORY ACTION - CALL THIS FIRST 🚨\n\nThis is STEP 2 of the 4-step workflow. You MUST call this action when you have both file ID and participants, BEFORE calling prepare_signature_request.\n\nWhat this does: Asks the user if they want security features (phone verification, password protection, Box login required).\n\nWhen to call: After you have file ID and participants, before prepare_signature_request.\n\nImportant: This action returns a message. Display it to the user and WAIT for their response. Do NOT call prepare_signature_request in the same message.",
    parameters: [
      {
        name: "fileId",
        type: "string",
        description: "The file ID for the signature request",
        required: true,
      },
      {
        name: "participantEmails",
        type: "string[]",
        description: "List of all participant email addresses",
        required: true,
      },
      {
        name: "requestSummary",
        type: "string",
        description: "Brief summary of request options if any (e.g., 'sequential, 30 days expiration, custom subject')",
        required: false,
      },
    ],
    handler: async ({ fileId, participantEmails, requestSummary }) => {
      const emails = participantEmails && Array.isArray(participantEmails) ? participantEmails : [];
      const emailList = emails.length > 0 ? emails.join(", ") : "the participants";
      const summary = requestSummary ? `\n\nRequest options: ${requestSummary}` : "";
      const gateKey = buildSecurityGateKey(
        String(fileId ?? "").trim(),
        emails
          .map((email) => String(email ?? "").trim())
          .filter(Boolean)
          .map((email) => ({ email, role: "signer" }))
      );
      securityGateRef.current.set(gateKey, {
        prompted: true,
        acknowledged: false,
        updatedAt: Date.now(),
      });
      securityGateRef.current.set(buildSecurityGateFileKey(String(fileId ?? "").trim()), {
        prompted: true,
        acknowledged: false,
        updatedAt: Date.now(),
      });
      
      console.log('[confirm_security_preferences] Called with:', { fileId, participantEmails, requestSummary });
      
      return `I have all the details for your signature request:\n- File ID: ${fileId}\n- Participants: ${emailList}${summary}\n\n**For added security**, I can require:\n• **Phone verification**: Participant enters a code sent to their phone\n• **Password protection**: You set a password they need to access the request\n• **Box login required**: Participant must be logged into their Box account\n\nWould you like any of these security features? You can specify which participant needs what security, or say "no" / "proceed without security" to continue.`;
    },
  });

  useCopilotAction({
    name: "prepare_signature_request",
    description:
      "⚠️ THIS IS STEP 3 - DO NOT CALL UNLESS YOU COMPLETED STEP 2 ⚠️\n\nSTEP 2 REQUIREMENT: Before calling this action, you MUST have already called confirm_security_preferences in a PREVIOUS message and received the user's response.\n\nWhat this does: Shows the user a document preview and request details (including security features in a table).\n\nWhen to call: ONLY after you called confirm_security_preferences AND the user responded to the security question.\n\nDO NOT CALL THIS if you haven't called confirm_security_preferences yet. You will skip the mandatory security check.",
    parameters: [
      { name: "fileId", type: "string", description: "Box file ID to be signed", required: true },
      { name: "parentFolderId", type: "string", description: "Box folder ID for signed document. Optional.", required: false },
      { name: "approverEmails", type: "string[]", description: "Emails of approvers.", required: false },
      { name: "signerEmails", type: "string[]", description: "Emails of signers.", required: false },
      { name: "finalCopyReaderEmails", type: "string[]", description: "Emails for final copy.", required: false },
      { name: "participants", type: "object[]", description: "Array of { email, role, verificationPhoneNumber?, password?, loginRequired? }. Use when adding security features per participant.", required: false },
      { name: "isSequential", type: "boolean", description: "Sequential signing order.", required: false },
      { name: "daysValid", type: "number", description: "Days until expiration.", required: false },
      { name: "areRemindersEnabled", type: "boolean", description: "Enable reminders.", required: false },
      { name: "name", type: "string", description: "Request name.", required: false },
      { name: "emailSubject", type: "string", description: "Email subject.", required: false },
      { name: "emailMessage", type: "string", description: "Email message body.", required: false },
      {
        name: "securityDecision",
        type: "string",
        description:
          "Human-in-the-loop security decision from the user: phone_verification, password_protection, box_login_required, multiple, or none.",
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
      
      // Log security feature usage
      const hasSecurityFeatures = participants.some(
        p => p.verificationPhoneNumber || p.password || p.loginRequired
      );
      console.log('[prepare_signature_request] Security features present:', hasSecurityFeatures);
      if (hasSecurityFeatures) {
        console.log('[prepare_signature_request] Security details:', 
          participants.map(p => ({
            email: p.email,
            phone: !!p.verificationPhoneNumber,
            password: !!p.password,
            loginRequired: !!p.loginRequired
          }))
        );
      }
      
      const fileId = params.fileId ? String(params.fileId).trim() : "";
      if (!fileId) throw new Error("fileId is required.");
      if (!participants.length)
        throw new Error("Specify who is involved: participants or approverEmails/signerEmails.");
      const gateKey = buildSecurityGateKey(fileId, participants);
      const fileGateKey = buildSecurityGateFileKey(fileId);
      const currentGate = securityGateRef.current.get(gateKey) ?? securityGateRef.current.get(fileGateKey);
      if (!currentGate?.prompted) {
        return "Before I can prepare this request, I must complete the security checkpoint first. Please confirm whether you want `phone_verification`, `password_protection`, `box_login_required`, `multiple`, or `none`.";
      }
      const securityDecision = normalizeSecurityDecision(params.securityDecision);
      if (!securityDecision) {
        return "Security confirmation required before preview. Please provide your choice: `phone_verification`, `password_protection`, `box_login_required`, `multiple`, or `none`.";
      }
      securityGateRef.current.set(gateKey, {
        prompted: true,
        acknowledged: true,
        decision: securityDecision,
        updatedAt: Date.now(),
      });
      securityGateRef.current.set(fileGateKey, {
        prompted: true,
        acknowledged: true,
        decision: securityDecision,
        updatedAt: Date.now(),
      });
      const participantsWithDecisionFallback = applySecurityDecisionFallback(
        participants,
        securityDecision
      );
      const missingSecurityDetails = getMissingSecurityDetailsMessage(
        participantsWithDecisionFallback,
        securityDecision
      );
      if (missingSecurityDetails) {
        return missingSecurityDetails;
      }
      const { fileName, requestSummary } = await stagePreviewAndDetails(
        params as Record<string, unknown>,
        participantsWithDecisionFallback
      );
      pendingCreateSignatureRef.current = buildCreateSignature(
        params as Record<string, unknown>,
        participantsWithDecisionFallback
      );
      return `I've prepared the signature request. The document **${fileName}** is shown in the preview panel next to this chat, with the request details below it. Please review and confirm: **Is this the correct document to sign?** Reply **Yes** or **Confirm** to send the signature request; I will then create it with the same parameters (${participants.length} participant(s), ${requestSummary.isSequential ? "sequential signing" : "any order"}${requestSummary.daysValid ? `, expires in ${requestSummary.daysValid} days` : ""}${requestSummary.areRemindersEnabled ? ", reminders enabled" : ""}).`;
    },
  });

  useCopilotAction({
    name: "create_signature_request",
    description:
      "THIS IS STEP 4 - Final step to actually create and send the signature request.\n\nCall this ONLY after: (1) You called confirm_security_preferences, (2) User responded to security question, (3) You called prepare_signature_request, (4) User confirmed the preview.\n\nRequired: fileId and at least one participant. For workflows with BOTH an approver and a signer: pass approverEmails and signerEmails. For security features (phone verification, password, Box login): use participants array with objects containing security properties. Order is always: approvers first, then signers, then final copy readers.",
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
          "Alternative: array of { email, role, verificationPhoneNumber?, password?, loginRequired? }. Use when adding security features per participant OR custom ordering. Order = sequence when isSequential.",
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
      {
        name: "securityDecision",
        type: "string",
        description:
          "Human-in-the-loop security decision already confirmed with the user: phone_verification, password_protection, box_login_required, multiple, or none.",
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
      const normalizedFileId = String(fileId).trim();
      const gateKey = buildSecurityGateKey(normalizedFileId, participants);
      const fileGateKey = buildSecurityGateFileKey(normalizedFileId);
      const gate = securityGateRef.current.get(gateKey) ?? securityGateRef.current.get(fileGateKey);
      const decisionFromParams = normalizeSecurityDecision(params.securityDecision);
      if (!gate?.prompted) {
        return "I can't send this request yet because the mandatory security checkpoint has not been completed. Please confirm whether you want phone verification, password protection, Box login required, or none.";
      }
      if (!gate.acknowledged && !decisionFromParams) {
        return "Before I can send this, I still need your explicit security decision: `phone_verification`, `password_protection`, `box_login_required`, `multiple`, or `none`.";
      }
      const effectiveDecision = decisionFromParams ?? gate.decision ?? "none";
      securityGateRef.current.set(gateKey, {
        prompted: true,
        acknowledged: true,
        decision: effectiveDecision,
        updatedAt: Date.now(),
      });
      securityGateRef.current.set(fileGateKey, {
        prompted: true,
        acknowledged: true,
        decision: effectiveDecision,
        updatedAt: Date.now(),
      });
      const participantsWithDecisionFallback = applySecurityDecisionFallback(
        participants,
        effectiveDecision
      );
      const missingSecurityDetails = getMissingSecurityDetailsMessage(
        participantsWithDecisionFallback,
        effectiveDecision
      );
      if (missingSecurityDetails) {
        return missingSecurityDetails;
      }

      const currentSignature = buildCreateSignature(
        params as Record<string, unknown>,
        participantsWithDecisionFallback
      );
      if (pendingCreateSignatureRef.current !== currentSignature) {
        const { fileName, requestSummary } = await stagePreviewAndDetails(
          params as Record<string, unknown>,
          participantsWithDecisionFallback
        );
        pendingCreateSignatureRef.current = currentSignature;
        return `Before sending, I must show the document preview and signing details. I have displayed **${fileName}** with all request details on screen. Please review and confirm in chat ("Yes" or "Confirm"), then ask me to send this exact request. Details: ${participantsWithDecisionFallback.length} participant(s), ${requestSummary.isSequential ? "sequential signing" : "any order"}${requestSummary.daysValid ? `, expires in ${requestSummary.daysValid} days` : ""}${requestSummary.areRemindersEnabled ? ", reminders enabled" : ""}.`;
      }

      const payload: Record<string, unknown> = {
        fileId: String(fileId),
        parentFolderId: parentFolderId != null && String(parentFolderId).trim() ? String(parentFolderId).trim() : undefined,
        participants: participantsWithDecisionFallback,
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
        const message = chatError(data, "Failed to create signature request.");
        console.error("[create_signature_request] Error:", message, data);
        return message;
      }
      const sr = data.signRequest;
      if (sr?.id) {
        lastCreatedSignRequestIdRef.current = sr.id;
        signRequestRoleCacheRef.current.set(
          sr.id,
          participantsWithDecisionFallback.map((p, index) => ({ email: p.email, role: p.role, order: !!isSequential ? index : undefined }))
        );
      }
      pendingCreateSignatureRef.current = null;
      securityGateRef.current.delete(gateKey);
      securityGateRef.current.delete(fileGateKey);
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
      if (data.error) return chatError(data, "List failed.");
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
        if (listData.error) return chatError(listData, "List failed.");
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
      if (r.error) return chatError(r, "Get failed.");
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
        if (listData.error) return chatError(listData, "List failed.");
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
      if (r.error) return chatError(r, "Get failed.");
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
      if (listData.error) return chatError(listData, "List failed.");

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

      // Optimistically remove matching requests from UI immediately
      const currentRequests = previewContext?.activeSignRequests ?? [];
      const idsToCancel = activeMatches.map(r => r.id);
      setActiveSignRequests(currentRequests.filter(r => !idsToCancel.includes(r.id)));
      
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

      // Refresh from server to ensure consistency
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
      
      // Optimistically remove from UI immediately
      const currentRequests = previewContext?.activeSignRequests ?? [];
      setActiveSignRequests(currentRequests.filter(r => r.id !== signRequestId));
      
      const data = await boxAction("cancel_signature_request", { signRequestId: String(signRequestId) });
      if (data.error) {
        // Restore on error
        setActiveSignRequests(currentRequests);
        return chatError(data, "Cancel failed.");
      }
      
      // Refresh from server to ensure consistency
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
      if (data.error) return chatError(data, "Resend failed.");
      return `Signature request \`${signRequestId}\` emails have been resent to outstanding signers.`;
    },
  });

  return null;
}
