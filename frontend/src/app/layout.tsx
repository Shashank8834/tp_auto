import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "TP Report Generator | Annexure 1 Automation",
  description: "Automated generation of Transfer Pricing Report (TNMM) - Annexure 1",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>
        <div className="app-container">
          {children}
        </div>
      </body>
    </html>
  );
}
