import {defineConfig,loadEnv} from "vite";
import react from "@vitejs/plugin-react";

// Hosting path is deployment configuration, not KERDOS business logic.
// Use "/" for a custom domain and "/repository-name/" for GitHub Pages.
export default defineConfig(({mode})=>{
  const env=loadEnv(mode,process.cwd(),"");
  const raw=env.VITE_BASE_PATH||"/";
  const base=raw==="/"?"/":`/${raw.replace(/^\/+|\/+$/g,"")}/`;
  return {base,plugins:[react()]};
});
