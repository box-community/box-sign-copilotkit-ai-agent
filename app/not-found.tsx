import Link from "next/link";

export default function NotFound() {
  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        padding: "2rem",
        fontFamily: "system-ui, sans-serif",
        textAlign: "center",
      }}
    >
      <h1 style={{ fontSize: "1.5rem", fontWeight: 600, marginBottom: "0.5rem" }}>
        404 — Page not found
      </h1>
      <p style={{ color: "var(--foreground)", opacity: 0.8, marginBottom: "1.5rem" }}>
        This page could not be found.
      </p>
      <Link
        href="/"
        style={{
          color: "var(--primary, #0061D5)",
          textDecoration: "underline",
          fontWeight: 500,
        }}
      >
        Back to Box Sign AI Assistant
      </Link>
    </div>
  );
}
