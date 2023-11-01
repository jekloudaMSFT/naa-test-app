import { IPublicClientApplication, PublicClientNext } from "@azure/msal-browser";

const msalConfig = {
    auth: {
        clientId: "bc4e0fa1-0903-4c5a-95e3-6587b1a21301",
        authority: "https://login.microsoftonline.com/common",
        supportsNestedAppAuth: true
    }
  };
  
  let pca: IPublicClientApplication;
  
  export function getNAAToken() {
    PublicClientNext.createPublicClientApplication(msalConfig).then((result) => {
        pca = result;
        ssoGetToken();
      });
  }
  
  export async function ssoGetToken(){
    //clientId: "a076930c-cfc9-4ebd-9607-7963bccbf666", //My msdn dev tenant
    //scope: "User.Read",
    //redirectUri: "https://localhost:3000",
    let activeAccount = pca.getActiveAccount();

    if (!activeAccount) {
        pca.loginPopup().then((result) => {
            if (result) {
                result.account && pca.setActiveAccount(result.account);
                activeAccount = result.account;
            }
        });
    }
  
    const tokenRequest = {
      scopes: ["User.Read"],
      account: activeAccount || undefined
    };
  
    pca.acquireTokenSilent(tokenRequest).then(async (result) => {
      console.log(result);
      const requestString = "https://graph.microsoft.com/v1.0/me";
      const headersInit = {'Authorization': result.accessToken};
      const requestInit = { 'headers': headersInit}
      if(requestString !== undefined){
          const result = await fetch(requestString, requestInit);
          if(result.ok){
            const data = await result.text();
            console.log(data);
          }else{
            //Handle whatever errors could happen that have nothing to do with identity
            console.log(result);
          }
  
      }else{
          //throw this should never happen
          throw new Error("unexpected: no requestString");
      }
    }).catch((error) => {
      console.log(error);
    });
  }