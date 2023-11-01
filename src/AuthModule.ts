import {
  IPublicClientApplication,
  PublicClientNext,
} from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "a076930c-cfc9-4ebd-9607-7963bccbf666",
    authority: "https://login.microsoftonline.com/common",
    supportsNestedAppAuth: true,
  },
};

class AuthModule {
  private msalClient?: IPublicClientApplication;

  init() {
    PublicClientNext.createPublicClientApplication(msalConfig).then(
      (client) => {
        this.msalClient = client;
      }
    );
  }

  getToken() {
    const activeAccount = this.msalClient?.getActiveAccount();
    const tokenRequest = {
      scopes: ["User.Read"],
      account: activeAccount || undefined,
    };

    this.msalClient
      ?.acquireTokenSilent(tokenRequest)
      .then(async (result) => {
        console.log(result);
        const requestString = "https://graph.microsoft.com/v1.0/me";
        const headersInit = { Authorization: result.accessToken };
        const requestInit = { headers: headersInit };
        if (requestString !== undefined) {
          const result = await fetch(requestString, requestInit);
          if (result.ok) {
            const data = await result.text();
            console.log(data);
          } else {
            //Handle whatever errors could happen that have nothing to do with identity
            console.log(result);
          }
        } else {
          //throw this should never happen
          throw new Error("unexpected: no requestString");
        }
      })
      .catch((error) => {
        console.log(error);
      });
  }
}
