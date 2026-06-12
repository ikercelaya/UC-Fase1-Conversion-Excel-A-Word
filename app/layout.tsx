import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Conversor Excel → Word · Universidad de Cantabria",
  description:
    "Genera informes Word a partir de archivos Excel — Universidad de Cantabria",
};

export default function RootLayout({
  children,
}: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
