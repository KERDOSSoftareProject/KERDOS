// KERDOS data adapter.
//
// This is the ONLY file in the application that knows which backend
// stores the data. Everything else — App.jsx, procurement.js,
// ingestion.js — talks to whatever this file exports and has no idea
// whether it is backed by Supabase, a self-hosted Postgres, a private
// API server, or on-device storage.
//
// Why this exists: KERDOS must not be permanently married to one
// vendor. Moving off Supabase should be a change to this one file, not
// a rewrite of 85 call sites scattered through the UI. The interface
// below (from/select/eq/order, auth, storage, channel) is the contract
// any replacement backend must implement.
//
// To move to a different backend later, write a module exporting the
// same shape and change the import in App.jsx. Nothing else changes.

import { createClient } from "@supabase/supabase-js";

// Config, not code. Set VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY at
// build time to point a build at a different project (a staging
// database, a self-hosted instance, a different customer's tenancy)
// without editing source.
//
// Worth being precise about what this does and does not buy: on a
// static host the built bundle still contains these values, so this is
// environment management, NOT secrecy. The anon key is designed to be
// public — the actual protection is row-level security on the database.
// Never put a service-role key here; that one is genuinely secret and
// belongs only on a server.
const SUPABASE_URL =
  import.meta.env?.VITE_SUPABASE_URL ||
  "https://antpbtorhqghrjqzftub.supabase.co";

const SUPABASE_ANON_KEY =
  import.meta.env?.VITE_SUPABASE_ANON_KEY ||
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFudHBidG9yaHFnaHJqcXpmdHViIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyOTMxMjMsImV4cCI6MjA5OTg2OTEyM30.vt4tdD7IaUKalzmpQrl5vD_hBb1lCOdu7D_LooC8qOQ";

export const data = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// Reports which backend a running build is actually pointed at, so a
// deployment can be identified without reading the bundle.
export const backendInfo = {
  kind: "supabase",
  url: SUPABASE_URL,
  configured: Boolean(import.meta.env?.VITE_SUPABASE_URL),
};
