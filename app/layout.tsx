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
  const description = "6 人扬州周末旅行共创：提交链接或直接写想法，交给 DeepSeek 整理并写入 Excel，再一起筛选住宿、美食、密室和周末行程。";
  return {
    metadataBase: base,
    title,
    description,
    icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
    openGraph: {
      title,
      description,
      type: "website",
      images: [{ url: new URL("/og-v2.png", base).toString(), width: 1731, height: 909, alt: "下扬州 · 6 人周末计划" }],
    },
    twitter: { card: "summary_large_image", title, description, images: [new URL("/og-v2.png", base).toString()] },
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
