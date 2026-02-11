import type { DocumentContext, DocumentInitialProps } from "next/document";
import Document, { Html, Head, Main, NextScript } from "next/document";

/**
 * Minimal _document so Next.js dev server does not throw
 * "Cannot find module '.next/server/pages/_document.js'" when handling
 * requests that hit the Pages Router path (e.g. /socket.io).
 * This app uses App Router only; this file exists only to satisfy the loader.
 */
export default class AppDocument extends Document {
  static async getInitialProps(
    ctx: DocumentContext
  ): Promise<DocumentInitialProps> {
    return Document.getInitialProps(ctx);
  }

  render() {
    return (
      <Html lang="en">
        <Head />
        <body>
          <Main />
          <NextScript />
        </body>
      </Html>
    );
  }
}
