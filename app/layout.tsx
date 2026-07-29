import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

export async function generateMetadata(): Promise<Metadata> {
  const incoming = await headers();
  const host = incoming.get("host") || "localhost:3100";
  const forwardedProtocol = incoming.get("x-forwarded-proto");
  const protocol = forwardedProtocol === "https" ? "https" : "http";
  const base = new URL(`${protocol}://${host}`);
  const title = "济州岛旅行共创台 · 六个人一起完成三日计划";
  const description = "团队提交济州岛旅行链接或直接写想法，交给 DeepSeek 整理并写入 Excel，再一起筛选住宿、美食、活动与最终行程。";
  return {
    metadataBase: base,
    title,
    description,
    icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
    openGraph: {
      title,
      description,
      type: "website",
      images: [{ url: new URL("/og.png", base).toString(), width: 1732, height: 908, alt: "出行共创台" }],
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
