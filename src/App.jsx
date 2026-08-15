import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  "https://antpbtorhqghrjqzftub.supabase.co",
  "sb_publishable_FQxGcfLMEo5K8iBLa7h2dA_eeaB7-W6"
);

export default function App() {
  return (
    <div style={{minHeight:"100vh",background:"#003584",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{color:"white",textAlign:"center"}}>
        <div style={{fontSize:60}}>🦉</div>
        <div style={{fontSize:24,fontWeight:900,letterSpacing:"0.2em",color:"#4A90D9",marginTop:8}}>KERDOS</div>
        <div style={{marginTop:12,opacity:0.7}}>Loading...</div>
      </div>
    </div>
  );
}
