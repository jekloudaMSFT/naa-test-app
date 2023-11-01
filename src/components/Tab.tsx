import { useContext, useState } from "react";
import { TeamsFxContext } from "./Context";
import { Button } from "@fluentui/react-components";
import { getNAAToken } from "../getNAAToken";

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  const [naaResult, setNaaResult] = useState("");
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
      <Button
        onClick={() => getNAAToken().then((result) => setNaaResult(result))}
      >
        Nested App Auth
      </Button>
      <div>{naaResult}</div>
    </div>
  );
}
