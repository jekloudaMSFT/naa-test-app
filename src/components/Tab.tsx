import { useContext } from "react";
import { Welcome } from "./sample/Welcome";
import { TeamsFxContext } from "./Context";
import config from "./sample/lib/config";
import { Button } from "@fluentui/react-components";
import { getNAAToken } from "../getNAAToken";

const showFunction = Boolean(config.apiName);

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  return (
    <div
      className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}
    >
      <Welcome showFunction={showFunction} />
      <Button onClick={() => getNAAToken()}>Nested App Auth</Button>
    </div>
  );
}
