import { NextRequest, NextResponse } from "next/server";
import { getBoxClient } from "@/lib/box-client";
import {
  FileBase,
  FolderMini,
  SignRequestCreateSigner,
} from "box-node-sdk/schemas";

type BoxAction =
  | "search_files"
  | "get_file_preview_info"
  | "create_signature_request"
  | "list_signature_requests"
  | "get_signature_request_status"
  | "cancel_signature_request"
  | "resend_signature_request";

type ParticipantRole = "signer" | "approver" | "final_copy_reader";
type RoleCacheEntry = {
  email: string;
  role: ParticipantRole;
  order?: number;
};

// Best-effort in-memory cache used to preserve participant roles for
// follow-up status/list calls when the upstream payload omits role.
const signRequestRoleCache = new Map<string, RoleCacheEntry[]>();

function normalizeEmail(email: string | undefined): string {
  return (email ?? "").trim().toLowerCase();
}

function toParticipantRole(role: unknown): ParticipantRole | undefined {
  if (role === "approver" || role === "signer" || role === "final_copy_reader") {
    return role;
  }
  return undefined;
}

function resolveRole(
  signRequestId: string | undefined,
  signer: { email?: string; order?: number; role?: string; rawData?: Record<string, unknown> }
): ParticipantRole | undefined {
  const explicit =
    toParticipantRole(signer.role) ??
    toParticipantRole((signer.rawData as { role?: unknown } | undefined)?.role);
  if (explicit) return explicit;

  const cached = signRequestId ? signRequestRoleCache.get(signRequestId) : undefined;
  if (!cached?.length) return undefined;

  const signerEmail = normalizeEmail(signer.email);
  if (signerEmail) {
    const byEmail = cached.find((c) => normalizeEmail(c.email) === signerEmail);
    if (byEmail) return byEmail.role;
  }
  if (typeof signer.order === "number") {
    const byOrder = cached.find((c) => c.order === signer.order);
    if (byOrder) return byOrder.role;
  }
  return undefined;
}

export async function POST(request: NextRequest) {
  let body: { action?: BoxAction; [key: string]: unknown };
  try {
    body = await request.json();
  } catch {
    return NextResponse.json(
      { error: "Invalid JSON body. Send { action, ...params }." },
      { status: 400 }
    );
  }
  const action = body?.action as BoxAction | undefined;
  if (!action || typeof action !== "string") {
    return NextResponse.json(
      {
        error:
          "Body must include action: search_files | get_file_preview_info | create_signature_request | list_signature_requests | get_signature_request_status | cancel_signature_request | resend_signature_request",
      },
      { status: 400 }
    );
  }

  try {
    switch (action) {
      case "search_files": {
        const q = (body.q ?? body.query ?? "").toString().trim();
        if (!q) {
          return NextResponse.json(
            { error: "Search requires q or query." },
            { status: 400 }
          );
        }
        const client = getBoxClient();
        const result = await client.search.searchForContent({
          query: q,
          type: "file",
          limit: 20,
        });
        const entries = result.entries ?? [];
        const files = entries
          .map(
            (item: {
              id?: string;
              name?: string;
              type?: string;
              parent?: { id?: string } | null;
            }) => ({
              id: item.id != null ? String(item.id) : "",
              name: item.name ?? undefined,
              type: item.type ?? "file",
              parentId:
                item.parent?.id != null ? String(item.parent.id) : undefined,
            })
          )
          .filter((f) => f.id);
        return NextResponse.json({ files });
      }

      case "get_file_preview_info": {
        const fileId = body.fileId?.toString()?.trim();
        if (!fileId) {
          return NextResponse.json(
            { error: "fileId is required for get_file_preview_info." },
            { status: 400 }
          );
        }
        const client = getBoxClient();
        try {
          const fileInfo = await client.files.getFileById(fileId, {
            queryParams: {
              fields: ["name", "expiring_embed_link"],
            },
          } as unknown as Parameters<typeof client.files.getFileById>[1]);
          const token = process.env.BOX_DEVELOPER_TOKEN;
          if (!token) {
            return NextResponse.json(
              { error: "BOX_DEVELOPER_TOKEN is not set." },
              { status: 500 }
            );
          }
          return NextResponse.json({
            fileId: fileInfo.id,
            fileName: fileInfo.name ?? "Document",
            token,
            embedUrl: fileInfo.expiringEmbedLink?.url,
          });
        } catch (fileErr: unknown) {
          const { status, message } = normalizeBoxError(fileErr);
          return NextResponse.json(
            {
              error:
                status === 404
                  ? `File not found (ID: ${fileId}).`
                  : message,
            },
            { status: status === 404 ? 404 : 500 }
          );
        }
      }

      case "list_signature_requests": {
        const client = getBoxClient();
        const result = await client.signRequests.getSignRequests({ limit: 25 });
        const entries = Array.from(result.entries ?? []).map((req) => ({
          id: req.id,
          status: req.status,
          name: req.name,
          createdAt: req.createdAt,
          signers: req.signers?.map((s) => ({
            email: s.email,
            role: resolveRole(req.id, s as { email?: string; order?: number; role?: string; rawData?: Record<string, unknown> }),
            order: s.order,
            status: (s as { signerDecision?: { type?: string } }).signerDecision?.type ?? "pending",
          })),
          sourceFileId: req.sourceFiles?.[0]?.id,
          parentFolder: req.parentFolder,
          daysValid: req.daysValid,
          areRemindersEnabled: req.areRemindersEnabled,
          autoExpireAt: req.autoExpireAt,
        }));
        return NextResponse.json({ signRequests: entries });
      }

      case "get_signature_request_status": {
        const id = (body.signRequestId ?? body.id)?.toString()?.trim();
        if (!id) {
          return NextResponse.json(
            { error: "signRequestId is required." },
            { status: 400 }
          );
        }
        const client = getBoxClient();
        const signRequest =
          await client.signRequests.getSignRequestById(id);
        return NextResponse.json({
          id: signRequest.id,
          status: signRequest.status,
          name: signRequest.name,
          signers: signRequest.signers?.map((s) => ({
            email: s.email,
            role: resolveRole(
              signRequest.id,
              s as { email?: string; order?: number; role?: string; rawData?: Record<string, unknown> }
            ),
            order: s.order,
            status: (s as { signerDecision?: { type?: string } }).signerDecision?.type ?? "pending",
          })),
          sourceFiles: signRequest.sourceFiles,
          parentFolder: signRequest.parentFolder,
          daysValid: signRequest.daysValid,
          areRemindersEnabled: signRequest.areRemindersEnabled,
          autoExpireAt: signRequest.autoExpireAt,
          prepareUrl: signRequest.prepareUrl,
        });
      }

      case "cancel_signature_request": {
        const id = (body.signRequestId ?? body.id)?.toString()?.trim();
        if (!id) {
          return NextResponse.json(
            { error: "signRequestId is required." },
            { status: 400 }
          );
        }
        const client = getBoxClient();
        await client.signRequests.cancelSignRequest(id);
        return NextResponse.json({ success: true, cancelled: id });
      }

      case "resend_signature_request": {
        const id = (body.signRequestId ?? body.id)?.toString()?.trim();
        if (!id) {
          return NextResponse.json(
            { error: "signRequestId is required." },
            { status: 400 }
          );
        }
        const client = getBoxClient();
        await client.signRequests.resendSignRequest(id);
        return NextResponse.json({ success: true, resend: id });
      }

      case "create_signature_request": {
        const fileId = body.fileId?.toString()?.trim() ?? "";
        const parentFolderId =
          body.parentFolderId != null ? String(body.parentFolderId).trim() : "";
        let participantsRaw = body.participants;
        if (typeof participantsRaw === "string") {
          try {
            participantsRaw = JSON.parse(participantsRaw) as unknown;
          } catch {
            participantsRaw = undefined;
          }
        }
        const isSequential = Boolean(body.isSequential);
        const daysValid =
          body.daysValid != null ? Number(body.daysValid) : undefined;
        const areRemindersEnabled = Boolean(body.areRemindersEnabled);
        const name = body.name != null ? String(body.name) : undefined;
        const emailSubject =
          body.emailSubject != null ? String(body.emailSubject) : undefined;
        const emailMessage =
          body.emailMessage != null ? String(body.emailMessage) : undefined;

        const validRoles: ParticipantRole[] = ["signer", "approver", "final_copy_reader"];
        let participantsList: Array<{ email?: string; role?: string }> = [];
        if (Array.isArray(participantsRaw)) {
          participantsList = participantsRaw;
        } else if (participantsRaw && typeof participantsRaw === "object" && !Array.isArray(participantsRaw)) {
          const obj = participantsRaw as Record<string, unknown>;
          if (typeof (obj as { email?: string }).email === "string") {
            participantsList = [participantsRaw as { email?: string; role?: string }];
          } else {
            participantsList = Object.keys(obj)
              .filter((k) => /^\d+$/.test(k))
              .sort((a, b) => Number(a) - Number(b))
              .map((k) => obj[k] as { email?: string; role?: string })
              .filter((p) => p && typeof p === "object");
          }
        }
        const getStr = (o: Record<string, unknown>, ...keys: string[]): string => {
          for (const k of keys) {
            if (o[k] != null && typeof o[k] === "string") return String(o[k]).trim();
          }
          return "";
        };
        const toList = (v: unknown): string[] =>
          Array.isArray(v) ? v.map((e) => String(e).trim()).filter(Boolean) : v != null ? [String(v).trim()].filter(Boolean) : [];
        const approverEmails = toList(body.approverEmails);
        const signerEmails = toList(body.signerEmails);
        const finalCopyReaderEmails = toList(body.finalCopyReaderEmails);
        const hasRoleSpecificLists = approverEmails.length > 0 || signerEmails.length > 0 || finalCopyReaderEmails.length > 0;

        let participants: Array<{ email: string; role: ParticipantRole }>;
        if (hasRoleSpecificLists) {
          participants = [
            ...approverEmails.map((email) => ({ email, role: "approver" as ParticipantRole })),
            ...signerEmails.map((email) => ({ email, role: "signer" as ParticipantRole })),
            ...finalCopyReaderEmails.map((email) => ({ email, role: "final_copy_reader" as ParticipantRole })),
          ];
        } else {
          participants = participantsList
            .map((p) => {
              const po = p && typeof p === "object" ? (p as Record<string, unknown>) : {};
              const email = getStr(po, "email", "Email").trim();
              const role = (getStr(po, "role", "Role") || "signer").toLowerCase() as ParticipantRole;
              if (!email) return null;
              return { email, role: validRoles.includes(role) ? role : "signer" };
            })
            .filter((p): p is { email: string; role: ParticipantRole } => p != null);
        }

        if (!fileId || !participants.length) {
          return NextResponse.json(
            {
              error: "fileId and participants are required. Pass participants (array of { email, role }), or approverEmails and signerEmails.",
              debug: {
                hint: "E.g. approverEmails: ['a@b.com'], signerEmails: ['c@d.com'] or participants: [{ email: 'a@b.com', role: 'approver' }, { email: 'c@d.com', role: 'signer' }].",
              },
            },
            { status: 400 }
          );
        }

        const placeholderDomains = ["example.com", "example.org", "example.net", "test.com", "test.org", "placeholder.com"];
        const allEmails = participants.map((p) => p.email);
        const invalidEmails = allEmails.filter((email) => {
          const domain = email.split("@")[1]?.toLowerCase();
          return domain && placeholderDomains.includes(domain);
        });
        if (invalidEmails.length > 0) {
          return NextResponse.json(
            {
              error: "Please use real email addresses. Placeholder emails (e.g. example@example.com) are not allowed.",
              debug: { hint: "Ask the user to provide actual participant email(s).", invalid: invalidEmails },
            },
            { status: 400 }
          );
        }

        const client = getBoxClient();
        let resolvedParentFolderId = parentFolderId || "";

        if (!resolvedParentFolderId || resolvedParentFolderId === "0") {
          try {
            const fileInfo = await client.files.getFileById(fileId);
            resolvedParentFolderId = fileInfo.parent?.id ?? "";
          } catch (fileErr: unknown) {
            const { status, message, requestId } = normalizeBoxError(fileErr);
            const error =
              status === 404
                ? `File not found (ID: ${fileId}). Use only file IDs from search_files; BOX_DEVELOPER_TOKEN must be for the Box user who owns the file.`
                : message;
            return NextResponse.json(
              {
                error,
                ...(status === 404 && {
                  debug: {
                    fileId,
                    requestId,
                    source: "get_file_by_id" as const,
                  },
                }),
              },
              { status }
            );
          }
          if (!resolvedParentFolderId || resolvedParentFolderId === "0") {
            return NextResponse.json(
              {
                error:
                  "The file is in the root folder. Box Sign requires a non-root folder. Please provide a parentFolderId.",
                reason: "ROOT_FOLDER",
                debug: {
                  hint: "Use the parentFolderId from search_files results, or create a folder in Box and use its ID.",
                  userMessage:
                    "The sign request was not created because the document is in your Box root folder. Box Sign requires the file to be inside a folder. To fix this: (1) Search for the file with search_files and use the parentFolderId from the results when creating the request, or (2) Create a folder in Box, put the file in that folder, and use that folder's ID as parentFolderId.",
                },
              },
              { status: 400 }
            );
          }
        }

        const signers: SignRequestCreateSigner[] = participants.map(
          (p: { email: string; role: ParticipantRole }, index: number) => ({
            email: p.email.trim(),
            role: p.role,
            order: isSequential ? index : undefined,
          })
        );

        const signRequest = await client.signRequests.createSignRequest({
          sourceFiles: [new FileBase({ id: fileId, type: "file" })],
          parentFolder: new FolderMini({
            id: resolvedParentFolderId,
            type: "folder",
          }),
          signers,
          ...(daysValid != null && { daysValid }),
          areRemindersEnabled,
          ...(name && { name }),
          ...(emailSubject && { emailSubject }),
          ...(emailMessage && { emailMessage }),
        });

        if (signRequest.id) {
          signRequestRoleCache.set(
            signRequest.id,
            participants.map((p, index) => ({
              email: p.email,
              role: p.role,
              order: isSequential ? index : undefined,
            }))
          );
        }

        return NextResponse.json({
          success: true,
          signRequest: {
            id: signRequest.id,
            status: signRequest.status,
            name: signRequest.name,
            prepareUrl: signRequest.prepareUrl,
            signers: signRequest.signers?.map((s) => ({
              email: s.email,
              role: resolveRole(
                signRequest.id,
                s as { email?: string; order?: number; role?: string; rawData?: Record<string, unknown> }
              ),
              order: s.order,
            })),
          },
        });
      }

      default:
        return NextResponse.json(
          { error: `Unknown action: ${action}` },
          { status: 400 }
        );
    }
  } catch (err: unknown) {
    const { status, message, requestId } = normalizeBoxError(err);
    const debug =
      process.env.NODE_ENV === "development" || status === 404
        ? {
            requestId,
            hint:
              status === 404 ? "Use file IDs from search_files." : undefined,
          }
        : undefined;
    return NextResponse.json(
      { error: message, ...(debug && { debug }) },
      { status }
    );
  }
}

function normalizeBoxError(
  err: unknown
): { status: number; message: string; requestId?: string } {
  const fallback = {
    status: 500,
    message: "Box API error.",
    requestId: undefined as string | undefined,
  };
  if (!err || typeof err !== "object") return fallback;

  const boxApi = err as {
    responseInfo?: {
      statusCode?: number;
      code?: string;
      requestId?: string;
      body?: Record<string, unknown> & { message?: string };
    };
    statusCode?: number;
    message?: string;
  };
  const statusCode =
    boxApi.responseInfo?.statusCode ?? boxApi.statusCode;
  const status =
    typeof statusCode === "number" && statusCode >= 400 && statusCode < 600
      ? statusCode
      : 500;
  const body = boxApi.responseInfo?.body;
  const bodyMessage =
    body && typeof body === "object" && "message" in body
      ? String((body as { message?: string }).message ?? "")
      : "";
  let message =
    boxApi.message ||
    bodyMessage ||
    (boxApi.responseInfo?.code ? `${boxApi.responseInfo.code}` : "") ||
    fallback.message;
  if (status === 404 && !message.toLowerCase().includes("not found")) {
    message = `Not found. ${message}`.trim();
  }
  const requestId = boxApi.responseInfo?.requestId;
  return { status, message, requestId };
}
