import {
  AccountInfo,
  IPublicClientApplication,
  PublicClientNext,
} from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "bc4e0fa1-0903-4c5a-95e3-6587b1a21301",
    authority: "https://login.microsoftonline.com/common",
    supportsNestedAppAuth: true,
  },
};

let pca: IPublicClientApplication;

export function initializePublicClient(): Promise<IPublicClientApplication> {
  console.log("Starting initializePublicClient");
  return PublicClientNext.createPublicClientApplication(msalConfig).then(
    (result) => {
      console.log("Client app created");
      pca = result;
      return pca;
    }
  );
}

export function getNAAToken(): Promise<string> {
  console.log("Starting getNAAToken");
  if (!pca) {
    return initializePublicClient().then((_client) => {
      return ssoGetToken();
    });
  } else {
    return ssoGetToken();
  }
}

export async function getActiveAccount(): Promise<AccountInfo> {
  console.log("Starting getActiveAccount");
  let activeAccount = null;
  try {
    console.log("getting active account");
    activeAccount = pca.getActiveAccount();
  } catch (error) {
    console.log(error);
  }

  if (!activeAccount) {
    console.log("No active account, trying login popup");
    await pca
      .loginPopup()
      .then((result) => {
        if (result) {
          result.account && pca.setActiveAccount(result.account);
          console.log(result);
          activeAccount = result.account;
        }
      })
  }
  if (activeAccount) {
    return activeAccount;
  } else {
    throw new Error("No active account found");
  }
}

export async function ssoGetToken(): Promise<string> {
  let activeAccount = await getActiveAccount();

  const tokenRequest = {
    scopes: ["User.Read"],
    account: activeAccount || undefined,
  };

  return pca
    .acquireTokenSilent(tokenRequest)
    .then((result) => {
      console.log(result);
      return result.accessToken;
    })
    .catch((error) => {
      console.log(error);
      // try to get token via popup
      return pca
        .acquireTokenPopup(tokenRequest)
        .then(async (result) => {
          console.log(result);
          return result.accessToken;
        })
        .catch((error) => {
          console.log(error);
          return JSON.stringify(error);
        });
    });
}

export async function ssoGetTokenAndFetchUser(): Promise<string> {
  return ssoGetToken().then((token) => {
    return fetchUserFromGraph(token);
  });
}

async function fetchUserFromGraph(accessToken: string): Promise<string> {
  const requestString = "https://graph.microsoft.com/v1.0/me";
  const headersInit = { Authorization: accessToken };
  const requestInit = { headers: headersInit };
  if (requestString !== undefined) {
    const result = await fetch(requestString, requestInit);
    if (result.ok) {
      const data = await result.text();
      console.log(data);
      return data;
    } else {
      //Handle whatever errors could happen that have nothing to do with identity
      console.log(result);
      return JSON.stringify(result);
    }
  } else {
    //throw this should never happen
    throw new Error("unexpected: no requestString");
  }
}
