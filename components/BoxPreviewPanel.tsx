"use client";

import { useEffect, useRef, useCallback, useId, useState } from "react";

const BOX_PREVIEW_SDK_VERSION = "2.106.0";
const LOCALE = "en-US";
const PREVIEW_CSS = `https://cdn01.boxcdn.net/platform/preview/${BOX_PREVIEW_SDK_VERSION}/${LOCALE}/preview.css`;
const PREVIEW_JS = `https://cdn01.boxcdn.net/platform/preview/${BOX_PREVIEW_SDK_VERSION}/${LOCALE}/preview.js`;

declare global {
  interface Window {
    Box?: {
      Preview: new () => {
        show: (fileId: string, token: string, options: { container: string; showDownload?: boolean; header?: string }) => void;
        hide: () => void;
      };
    };
  }
}

export type RequestSummary = {
  fileId: string;
  fileName: string;
  participants: Array<{ email: string; role: string }>;
  isSequential?: boolean;
  daysValid?: number;
  areRemindersEnabled?: boolean;
  name?: string;
  emailSubject?: string;
};

export type PreviewData = {
  fileId: string;
  fileName: string;
  token: string;
  embedUrl?: string;
  requestSummary: RequestSummary;
} | {
  documents: Array<{
    fileId: string;
    fileName: string;
    token: string;
    embedUrl?: string;
    requestSummary: RequestSummary;
  }>;
};

function loadScript(src: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) {
      resolve();
      return;
    }
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(script);
  });
}

function loadStylesheet(href: string): Promise<void> {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`link[href="${href}"]`)) {
      resolve();
      return;
    }
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = href;
    link.onload = () => resolve();
    link.onerror = () => reject(new Error(`Failed to load ${href}`));
    document.head.appendChild(link);
  });
}

function roleLabel(role: string): string {
  switch (role) {
    case "approver":
      return "Approver";
    case "final_copy_reader":
      return "Gets a copy";
    case "signer":
    default:
      return "Signer";
  }
}

function getEffectiveReminderDays(daysValid?: number): number[] {
  const cadence = [3, 8, 13, 18];
  if (daysValid == null || daysValid <= 0) return cadence;
  return cadence.filter((day) => day <= daysValid);
}

export function BoxPreviewPanel({
  data,
  onDismiss,
}: {
  data: PreviewData;
  onDismiss?: () => void;
}) {
  const containerId = useId().replace(/:/g, "");
  const containerRef = useRef<HTMLDivElement>(null);
  const previewInstanceRef = useRef<InstanceType<NonNullable<typeof window.Box>["Preview"]> | null>(null);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [iframeLoaded, setIframeLoaded] = useState(false);
  const [currentIndex, setCurrentIndex] = useState(0);
  
  // Check if data contains multiple documents
  const isMultiDocument = "documents" in data;
  const documents = isMultiDocument ? data.documents : [data];
  const currentDoc = documents[currentIndex] || documents[0];
  const totalDocs = documents.length;

  // Reset iframe loaded state when document changes
  useEffect(() => {
    setIframeLoaded(false);
    setPreviewError(null);
  }, [currentIndex]);

  const showPreview = useCallback(
    async (fileId: string, token: string) => {
      if (!containerRef.current || typeof window === "undefined") return;
      setPreviewError(null);
      try {
        await loadStylesheet(PREVIEW_CSS);
        await loadScript(PREVIEW_JS);
      } catch (err) {
        console.error("[BoxPreviewPanel] Failed to load Box Preview assets:", err);
        setPreviewError("Could not load Box Preview assets. Please refresh the page and try again.");
        return;
      }
      const Box = window.Box;
      if (!Box?.Preview) {
        console.error("[BoxPreviewPanel] Box.Preview not available");
        setPreviewError("Box Preview SDK is not available in the browser.");
        return;
      }
      if (previewInstanceRef.current) {
        try {
          previewInstanceRef.current.hide();
        } catch (_) {}
        previewInstanceRef.current = null;
      }
      const preview = new Box.Preview();
      previewInstanceRef.current = preview;
      try {
        preview.show(fileId, token, {
          container: `#${containerId}`,
          showDownload: false,
          header: "light",
        });
      } catch (err) {
        console.error("[BoxPreviewPanel] Failed to render preview:", err);
        setPreviewError("Failed to render the document preview. Please verify file access and token.");
      }
    },
    [containerId]
  );

  useEffect(() => {
    if (currentDoc?.embedUrl) return;
    if (currentDoc?.fileId && currentDoc?.token) {
      showPreview(currentDoc.fileId, currentDoc.token);
    }
    return () => {
      if (previewInstanceRef.current) {
        try {
          previewInstanceRef.current.hide();
        } catch (_) {}
        previewInstanceRef.current = null;
      }
    };
  }, [currentDoc?.fileId, currentDoc?.token, showPreview, currentIndex]);

  if (!currentDoc) return null;

  const { fileName, requestSummary } = currentDoc;
  const effectiveReminderDays = getEffectiveReminderDays(requestSummary.daysValid);
  const reminderLabel =
    effectiveReminderDays.length > 0
      ? `Reminders: enabled (days ${effectiveReminderDays.join(", ")})`
      : "Reminders: enabled (no reminder day occurs before expiration)";

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        minHeight: "calc(100vh - 2rem)",
        width: "100%",
        background: "white",
        border: "var(--border-default)",
        borderRadius: "var(--radius-xl)",
        overflow: "hidden",
        boxShadow: "var(--shadow-lg)",
      }}
    >
      <div
        style={{
          padding: "var(--space-lg) var(--space-xl)",
          borderBottom: "var(--border-default)",
          background: "linear-gradient(135deg, #0061D5 0%, #003D8F 100%)",
          color: "white",
          borderTopLeftRadius: "var(--radius-xl)",
          borderTopRightRadius: "var(--radius-xl)",
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "var(--space-sm)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "var(--space-md)" }}>
            <h2 style={{ margin: 0, fontSize: "var(--text-xl)", fontWeight: "var(--font-semibold)", letterSpacing: "-0.0125em" }}>Document Preview</h2>
            {totalDocs > 1 && (
              <span style={{ 
                fontSize: "var(--text-sm)", 
                fontWeight: "var(--font-medium)", 
                background: "rgba(255, 255, 255, 0.2)",
                padding: "0.25rem 0.75rem",
                borderRadius: "var(--radius-md)",
                border: "1px solid rgba(255, 255, 255, 0.3)"
              }}>
                {currentIndex + 1} of {totalDocs}
              </span>
            )}
          </div>
          {onDismiss && (
            <button
              type="button"
              onClick={onDismiss}
              style={{
                padding: "0.5rem 1rem",
                fontSize: "var(--text-sm)",
                fontWeight: "var(--font-medium)",
                background: "rgba(255, 255, 255, 0.15)",
                border: "1px solid rgba(255, 255, 255, 0.3)",
                borderRadius: "var(--radius-lg)",
                color: "white",
                cursor: "pointer",
                transition: "all 150ms ease",
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.background = "rgba(255, 255, 255, 0.25)";
                e.currentTarget.style.borderColor = "rgba(255, 255, 255, 0.5)";
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = "rgba(255, 255, 255, 0.15)";
                e.currentTarget.style.borderColor = "rgba(255, 255, 255, 0.3)";
              }}
            >
              Close Preview
            </button>
          )}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "var(--space-md)", marginTop: "var(--space-xs)" }}>
          <p style={{ margin: 0, fontSize: "var(--text-base)", opacity: 0.95, fontWeight: "var(--font-medium)", flex: 1 }}>{fileName}</p>
          {totalDocs > 1 && (
            <div style={{ display: "flex", gap: "var(--space-xs)" }}>
              <button
                type="button"
                onClick={() => setCurrentIndex(Math.max(0, currentIndex - 1))}
                disabled={currentIndex === 0}
                style={{
                  padding: "0.5rem 1rem",
                  fontSize: "var(--text-sm)",
                  fontWeight: "var(--font-medium)",
                  background: currentIndex === 0 ? "rgba(255, 255, 255, 0.1)" : "rgba(255, 255, 255, 0.15)",
                  border: "1px solid rgba(255, 255, 255, 0.3)",
                  borderRadius: "var(--radius-md)",
                  color: "white",
                  cursor: currentIndex === 0 ? "not-allowed" : "pointer",
                  opacity: currentIndex === 0 ? 0.5 : 1,
                  transition: "all 150ms ease",
                }}
                onMouseEnter={(e) => {
                  if (currentIndex !== 0) {
                    e.currentTarget.style.background = "rgba(255, 255, 255, 0.25)";
                  }
                }}
                onMouseLeave={(e) => {
                  if (currentIndex !== 0) {
                    e.currentTarget.style.background = "rgba(255, 255, 255, 0.15)";
                  }
                }}
              >
                ← Previous
              </button>
              <button
                type="button"
                onClick={() => setCurrentIndex(Math.min(totalDocs - 1, currentIndex + 1))}
                disabled={currentIndex === totalDocs - 1}
                style={{
                  padding: "0.5rem 1rem",
                  fontSize: "var(--text-sm)",
                  fontWeight: "var(--font-medium)",
                  background: currentIndex === totalDocs - 1 ? "rgba(255, 255, 255, 0.1)" : "rgba(255, 255, 255, 0.15)",
                  border: "1px solid rgba(255, 255, 255, 0.3)",
                  borderRadius: "var(--radius-md)",
                  color: "white",
                  cursor: currentIndex === totalDocs - 1 ? "not-allowed" : "pointer",
                  opacity: currentIndex === totalDocs - 1 ? 0.5 : 1,
                  transition: "all 150ms ease",
                }}
                onMouseEnter={(e) => {
                  if (currentIndex !== totalDocs - 1) {
                    e.currentTarget.style.background = "rgba(255, 255, 255, 0.25)";
                  }
                }}
                onMouseLeave={(e) => {
                  if (currentIndex !== totalDocs - 1) {
                    e.currentTarget.style.background = "rgba(255, 255, 255, 0.15)";
                  }
                }}
              >
                Next →
              </button>
            </div>
          )}
        </div>
      </div>
      <div
        style={{
          flex: "0 0 auto",
          padding: "1.75rem 2rem",
          borderBottom: "1px solid rgba(0, 0, 0, 0.15)",
          background: "white",
          fontSize: "var(--text-base)",
        }}
      >
        <div style={{ 
          fontWeight: 700, 
          marginBottom: "1.25rem", 
          color: "#0F172A", 
          fontSize: "1.125rem", 
          letterSpacing: "-0.0125em" 
        }}>
          Request Details
        </div>
        <div style={{ 
          display: "grid", 
          gap: "0.875rem", 
          margin: 0, 
          lineHeight: 1.6, 
          color: "#334155" 
        }}>
          {requestSummary.participants.map((p, i) => (
            <div key={i} style={{ display: "flex", alignItems: "baseline", gap: "0.75rem" }}>
              <span style={{ 
                fontWeight: 600, 
                minWidth: "130px",
                color: "#0061D5",
                fontSize: "0.9375rem" 
              }}>
                {roleLabel(p.role)}:
              </span>
              <span style={{ fontSize: "0.9375rem", color: "#1E293B", fontWeight: 500 }}>{p.email}</span>
            </div>
          ))}
          {requestSummary.isSequential && (
            <div style={{ display: "flex", alignItems: "baseline", gap: "0.75rem" }}>
              <span style={{ fontWeight: 600, minWidth: "130px", fontSize: "0.9375rem", color: "#475569" }}>Signing order:</span>
              <span style={{ fontSize: "0.9375rem", color: "#1E293B", fontWeight: 500 }}>Sequential</span>
            </div>
          )}
          {requestSummary.daysValid != null && requestSummary.daysValid > 0 && (
            <div style={{ display: "flex", alignItems: "baseline", gap: "0.75rem" }}>
              <span style={{ fontWeight: 600, minWidth: "130px", fontSize: "0.9375rem", color: "#475569" }}>Expires in:</span>
              <span style={{ fontSize: "0.9375rem", color: "#1E293B", fontWeight: 500 }}>{requestSummary.daysValid} days</span>
            </div>
          )}
          {requestSummary.areRemindersEnabled && (
            <div style={{ display: "flex", alignItems: "baseline", gap: "0.75rem" }}>
              <span style={{ fontWeight: 600, minWidth: "130px", fontSize: "0.9375rem", color: "#475569" }}>Reminders:</span>
              <span style={{ fontSize: "0.9375rem", color: "#1E293B", fontWeight: 500 }}>
                Enabled {effectiveReminderDays.length > 0 ? `(days ${effectiveReminderDays.join(", ")})` : "(none before expiration)"}
              </span>
            </div>
          )}
          {requestSummary.name && (
            <div style={{ display: "flex", alignItems: "baseline", gap: "0.75rem" }}>
              <span style={{ fontWeight: 600, minWidth: "130px", fontSize: "0.9375rem", color: "#475569" }}>Request name:</span>
              <span style={{ fontSize: "0.9375rem", color: "#1E293B", fontWeight: 500 }}>{requestSummary.name}</span>
            </div>
          )}
          {requestSummary.emailSubject && (
            <div style={{ display: "flex", alignItems: "baseline", gap: "0.75rem" }}>
              <span style={{ fontWeight: 600, minWidth: "130px", fontSize: "0.9375rem", color: "#475569" }}>Email subject:</span>
              <span style={{ fontSize: "0.9375rem", color: "#1E293B", fontWeight: 500 }}>{requestSummary.emailSubject}</span>
            </div>
          )}
        </div>
        <div
          style={{
            marginTop: "1.5rem",
            padding: "1.25rem 1.5rem",
            background: "rgba(0, 97, 213, 0.08)",
            borderRadius: "var(--radius-lg)",
            border: "1px solid rgba(0, 97, 213, 0.25)",
            fontWeight: 500,
            color: "#1E293B",
            fontSize: "0.9375rem",
            lineHeight: "1.7",
            boxShadow: "0 1px 2px rgba(0, 97, 213, 0.1)",
          }}
        >
          <span style={{ color: "#334155" }}>
            {totalDocs > 1 
              ? `Review all ${totalDocs} documents using the navigation buttons above. When ready, reply ` 
              : "Is this the correct document to sign? Reply "}
          </span>
          <strong style={{ color: "#0061D5", fontWeight: 700 }}>Yes</strong>
          <span style={{ color: "#334155" }}> or </span>
          <strong style={{ color: "#0061D5", fontWeight: 700 }}>Confirm</strong>
          <span style={{ color: "#334155" }}> in the chat to send the signature request{totalDocs > 1 ? "s" : ""}.</span>
        </div>
      </div>
      <div
        id={containerId}
        className="box-preview-container"
        ref={containerRef}
        style={{
          flex: "1 1 auto",
          minHeight: 520,
          width: "100%",
          position: "relative",
          background: "#fff",
          borderBottomLeftRadius: "var(--radius-xl)",
          borderBottomRightRadius: "var(--radius-xl)",
        }}
      >
        <div
          style={{
            position: "absolute",
            right: 16,
            top: 16,
            zIndex: 2,
            display: "flex",
            gap: 8,
          }}
        >
          <a
            href={currentDoc.embedUrl || `https://app.box.com/file/${currentDoc.fileId}`}
            target="_blank"
            rel="noreferrer"
            style={{
              fontSize: "0.875rem",
              fontWeight: 500,
              textDecoration: "none",
              color: "#0061D5",
              background: "rgba(255, 255, 255, 0.95)",
              border: "1px solid rgba(0, 97, 213, 0.2)",
              borderRadius: "var(--radius-md)",
              padding: "0.5rem 0.875rem",
              boxShadow: "0 2px 4px rgba(0, 0, 0, 0.08)",
              transition: "all 150ms ease",
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.background = "white";
              e.currentTarget.style.borderColor = "#0061D5";
              e.currentTarget.style.boxShadow = "0 4px 6px rgba(0, 0, 0, 0.12)";
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.background = "rgba(255, 255, 255, 0.95)";
              e.currentTarget.style.borderColor = "rgba(0, 97, 213, 0.2)";
              e.currentTarget.style.boxShadow = "0 2px 4px rgba(0, 0, 0, 0.08)";
            }}
          >
            Open in Box
          </a>
        </div>
        {currentDoc.embedUrl ? (
          <iframe
            src={currentDoc.embedUrl}
            key={currentDoc.fileId}
            title={`Preview of ${fileName}`}
            onLoad={() => setIframeLoaded(true)}
            onError={() =>
              setPreviewError(
                "Unable to render iframe preview for this document in-app. Use 'Open in Box' while we keep the signing details on-screen."
              )
            }
            style={{
              width: "100%",
              height: "100%",
              minHeight: 520,
              border: "none",
              display: "block",
            }}
            allow="fullscreen"
          />
        ) : null}
        {currentDoc.embedUrl && !iframeLoaded && !previewError ? (
          <div
            style={{
              position: "absolute",
              left: 16,
              top: 16,
              zIndex: 2,
              padding: "0.5rem 0.875rem",
              borderRadius: "var(--radius-md)",
              background: "rgba(255, 255, 255, 0.95)",
              border: "1px solid rgba(0, 0, 0, 0.12)",
              fontSize: "0.875rem",
              fontWeight: 500,
              color: "#334155",
              boxShadow: "0 2px 4px rgba(0, 0, 0, 0.08)",
            }}
          >
            Loading document preview...
          </div>
        ) : null}
        {previewError ? (
          <div
            style={{
              margin: "1.25rem",
              padding: "1rem 1.25rem",
              borderRadius: "var(--radius-lg)",
              border: "1px solid rgba(180, 35, 24, 0.25)",
              background: "rgba(180, 35, 24, 0.08)",
              color: "#991B1B",
              fontSize: "0.9375rem",
              lineHeight: "1.6",
              fontWeight: 500,
            }}
          >
            {previewError}
          </div>
        ) : null}
      </div>
    </div>
  );
}
