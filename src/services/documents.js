// Provider-neutral document operations used by KERDOS screens.
// Paths are KERDOS conventions; storage mechanics belong to the adapter.
export function createDocumentService(backend){
  return {
    async uploadOriginal(organizationId,vendorId,file){
      const path=`${organizationId}/${vendorId}/${Date.now()}_${file.name}`;
      try{await backend.documents.upload(path,file);return {path,name:file.name,error:null};}
      catch(error){return {path:null,name:null,error};}
    },
    async uploadLogo(organizationId,file){
      const path=`${organizationId}/logo/${Date.now()}_${file.name}`;
      try{await backend.documents.upload(path,file,{upsert:true});return {path,error:null};}
      catch(error){return {path:null,error};}
    },
    async signedUrl(path,seconds=3600){
      if(!path)return null;
      return (await backend.documents.signedUrl(path,seconds)).signedUrl;
    },
    remove(paths){return backend.documents.remove(paths);},
  };
}
