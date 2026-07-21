import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

export async function generateMetadata(): Promise<Metadata> {
  const incoming = await headers();
  const host = incoming.get("host") || "localhost:3000";
  const forwardedProtocol = incoming.get("x-forwarded-proto");
  const protocol = forwardedProtocol === "https" ? "https" : "http";
  const base = new URL(`${protocol}://${host}`);
  const title = "下扬州 · 团队旅行共创台";
  const description = "6 人扬州周末旅行共创：收集链接、DeepSeek 整理、Excel 核对，并安排两晚住宿、烧烤、早餐和密室。";
  return {
    metadataBase: base,
    title,
    description,
    icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
    openGraph: {
      title,
      description,
      type: "website",
      images: [{ url: new URL("/og.png", base).toString(), width: 1736, height: 909, alt: "下扬州 · 团队旅行共创台" }],
    },
    twitter: { card: "summary_large_image", title, description, images: [new URL("/og.png", base).toString()] },
  };
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
