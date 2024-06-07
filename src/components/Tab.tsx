import { useContext } from "react";
import { TeamsFxContext } from "./Context";
import { getActiveAccount, getNAAToken, initializePublicClient, getTokenAndFetchUser } from "../getNAAToken";
import ApiControl from "./ApiControl";
import { nestedAppAuth } from "@microsoft/teams-js";

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  return (
    <div
      className={
        themeString === "default"
          ? "light"
          : themeString === "dark"
          ? "dark"
          : "contrast"
      }
    >
      <div>2024060501</div>
      <div className="api-container">
        <ApiControl
          apiName="isNAAChannelRecommended"
          onClick={async () => JSON.stringify(await nestedAppAuth.isNAAChannelRecommended())}
        />
        <ApiControl
          apiName="Initialize public client application"
          onClick={async () => JSON.stringify(await initializePublicClient())}
        />
        <ApiControl
          apiName="Get auth token"
          onClick={async () => await getNAAToken()}
        />
        <ApiControl
          apiName="Get active account"
          onClick={async () => JSON.stringify(await getActiveAccount())}
        />
        <ApiControl
          apiName="Get user info from Graph"
          onClick={async () => await getTokenAndFetchUser()}
        />
      </div>
    </div>
  );
}
