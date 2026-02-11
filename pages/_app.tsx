import type { AppProps } from "next/app";

/**
 * Minimal _app so Next.js dev server can resolve pages when needed.
 * This app uses App Router only; this file exists to satisfy the loader.
 */
export default function App({ Component, pageProps }: AppProps) {
  return <Component {...pageProps} />;
}
