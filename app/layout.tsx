import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Box Sign AI Assistant",
  description: "Natural language e-signature flow with Box and CopilotKit",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased min-h-screen">{children}</body>
    </html>
  );
}
