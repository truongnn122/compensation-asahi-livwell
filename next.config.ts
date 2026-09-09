import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // firebase-admin's auth module pulls in jwks-rsa -> jose, and jose is
  // ESM-only while jwks-rsa still requires() it — bundling firebase-admin
  // breaks that interop. Externalizing it lets Node's own module resolution
  // handle it at runtime instead of the bundler rewriting it.
  serverExternalPackages: ["firebase-admin"],
  // Firebase client SDK config is public by design, but we don't use the
  // NEXT_PUBLIC_ prefix — this explicitly inlines these into the browser
  // bundle at build time instead.
  env: {
    FIREBASE_API_KEY: process.env.FIREBASE_API_KEY,
    FIREBASE_AUTH_DOMAIN: process.env.FIREBASE_AUTH_DOMAIN,
    FIREBASE_PROJECT_ID: process.env.FIREBASE_PROJECT_ID,
    FIREBASE_STORAGE_BUCKET: process.env.FIREBASE_STORAGE_BUCKET,
    FIREBASE_MESSAGING_SENDER_ID: process.env.FIREBASE_MESSAGING_SENDER_ID,
    FIREBASE_APP_ID: process.env.FIREBASE_APP_ID,
    FIRESTORE_DATABASE_ID: process.env.FIRESTORE_DATABASE_ID,
  },
  experimental: {
    serverActions: {
      bodySizeLimit: "10mb",
    },
  },
};

export default nextConfig;
