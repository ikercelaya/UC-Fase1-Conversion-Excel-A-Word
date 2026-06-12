import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Evita que Next confunda la raíz del proyecto si existen otros lockfiles en el sistema.
  outputFileTracingRoot: process.cwd(),
};

export default nextConfig;
