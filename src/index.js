import { assertBackendContract } from "./contract.js";
import { createSupabaseBackend } from "./supabase.js";

const runtime=globalThis.__KERDOS_CONFIG__ || {};
const env=import.meta.env || {};
const provider=runtime.backendProvider || env.VITE_KERDOS_BACKEND || "supabase";

function unconfigured(reason){
  const fail=async()=>{throw new Error(reason);};
  return {
    kind:"unconfigured",
    session:{get:async()=>null,subscribe:()=>()=>{},signIn:fail,signUp:fail,signOut:fail},
    workspace:{memberships:fail,snapshot:fail},
    documents:{upload:fail,signedUrl:fail,remove:fail},
    realtime:{subscribeToOrganization:()=>()=>{}},
    pricing:{applyQuote:fail},
    invoices:{record:fail},
    team:{acceptInvite:fail},
    commands:{table:()=>{throw new Error(reason);}},
  };
}

const factories={
  supabase:()=>createSupabaseBackend({
    url:runtime.backendUrl || env.VITE_SUPABASE_URL,
    anonKey:runtime.backendPublicKey || env.VITE_SUPABASE_ANON_KEY,
  }),
};

const hasSupabaseConfig=!!(runtime.backendUrl||env.VITE_SUPABASE_URL)&&!!(runtime.backendPublicKey||env.VITE_SUPABASE_ANON_KEY);
const selected=!factories[provider]
  ?unconfigured(`Unsupported KERDOS backend provider: ${provider}`)
  :provider==="supabase"&&!hasSupabaseConfig
    ?unconfigured("KERDOS backend is not configured for this deployment")
    :factories[provider]();
export const backend=assertBackendContract(selected);
export const backendInfo=Object.freeze({kind:backend.kind,configured:backend.kind!=="unconfigured"});
