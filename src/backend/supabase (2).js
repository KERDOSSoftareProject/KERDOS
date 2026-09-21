import { createClient } from "@supabase/supabase-js";
import { BackendError, providerResult } from "./contract.js";

const SNAPSHOT_QUERIES = Object.freeze([
  ["vendors", c=>c.from("vendors").select("*").eq("is_active",true).order("name")],
  ["catalogItems", c=>c.from("catalog_items").select("*,catalog_categories(name)").order("master_item_number")],
  ["categories", c=>c.from("catalog_categories").select("*").order("name")],
  ["vendorItems", c=>c.from("vendor_items").select("*")],
  ["mappings", c=>c.from("item_mappings").select("*")],
  ["invoices", c=>c.from("invoices").select("*,vendors(name),invoice_lines(*)").order("invoice_date",{ascending:false,nullsFirst:false}).order("created_at",{ascending:false}).limit(30)],
  ["purchaseOrders", c=>c.from("purchase_orders").select("*,purchase_order_lines(*)").order("created_at",{ascending:false}).limit(30)],
  ["priceHistory", c=>c.from("price_history").select("*").order("effective_date",{ascending:false}).order("id",{ascending:false}).limit(2000)],
  ["vocabulary", c=>c.from("org_vocabulary").select("*")],
]);

export function createSupabaseBackend({url,anonKey}) {
  if(!url||!anonKey) throw new Error("Supabase adapter requires an external URL and anonymous key");
  const client=createClient(url,anonKey);
  return {
    kind:"supabase",
    session:{
      get:()=>providerResult(client.auth.getSession(),"Load session").then(x=>x?.session||null),
      subscribe(listener){
        const result=client.auth.onAuthStateChange((event,session)=>listener({event,session:session||null}));
        return ()=>result?.data?.subscription?.unsubscribe?.();
      },
      signIn:(credentials)=>providerResult(client.auth.signInWithPassword(credentials),"Sign in"),
      signUp:(credentials)=>providerResult(client.auth.signUp(credentials),"Create account"),
      signOut:()=>providerResult(client.auth.signOut(),"Sign out"),
    },
    workspace:{
      memberships(userId){
        return providerResult(client.from("organization_members").select("organization_id,role,organizations(*)").eq("user_id",userId),"Load organizations");
      },
      async snapshot(organizationId){
        const entries=await Promise.all(SNAPSHOT_QUERIES.map(async ([name,build])=>{
          const query=build(client).eq("organization_id",organizationId);
          return [name,await providerResult(query,`Load ${name}`) || []];
        }));
        return Object.fromEntries(entries);
      },
    },
    documents:{
      upload(path,file,options){return providerResult(client.storage.from("documents").upload(path,file,options),"Upload document");},
      signedUrl(path,seconds=3600){return providerResult(client.storage.from("documents").createSignedUrl(path,seconds),"Open document");},
      remove(paths){return providerResult(client.storage.from("documents").remove(paths),"Remove document");},
    },
    realtime:{
      subscribeToOrganization(organizationId,onChange){
        const channel=client.channel(`kerdos:${organizationId}`)
          .on("postgres_changes",{event:"*",schema:"public",table:"vendor_items",filter:`organization_id=eq.${organizationId}`},onChange)
          .on("postgres_changes",{event:"*",schema:"public",table:"invoices",filter:`organization_id=eq.${organizationId}`},onChange)
          .subscribe();
        return ()=>client.removeChannel(channel);
      },
    },
    pricing:{
      applyQuote(quote){
        return providerResult(client.rpc("kerdos_apply_price_quote",{
          p_vendor_item_id:quote.vendorItemId||null,
          p_organization_id:quote.organizationId,
          p_vendor_id:quote.vendorId,
          p_vendor_item_code:quote.vendorItemCode||null,
          p_description:quote.description,
          p_pack_size:quote.packSize||null,
          p_price:quote.price??null,
          p_price_unavailable:!!quote.priceUnavailable,
          p_effective_date:quote.effectiveDate,
          p_quote_valid_until:quote.quoteValidUntil||null,
          p_source_file_path:quote.sourceFilePath||null,
          p_source_file_name:quote.sourceFileName||null,
          p_source_line:quote.sourceLine||null,
          p_source_document_id:quote.sourceDocumentId,
        }),"Apply price quotation");
      },
    },
    invoices:{
      record(header,lines){
        return providerResult(client.rpc("kerdos_record_invoice",{p_header:header,p_lines:lines}),"Record invoice");
      },
    },
    team:{
      async acceptInvite(code,userId){
        const rows=await providerResult(client.rpc("kerdos_accept_invite",{p_code:code,p_user_id:userId}),"Accept invitation");
        const accepted=rows?.[0];
        if(!accepted) throw new BackendError("Accept invitation",new Error("Invitation was not accepted"));
        return accepted;
      },
    },
    // Temporary migration bridge. Each call site moved behind a KERDOS
    // service shrinks this surface; new feature code must not use it.
    commands:{table:name=>client.from(name)},
  };
}
