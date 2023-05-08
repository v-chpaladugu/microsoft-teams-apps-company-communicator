import { teamsDarkTheme, teamsHighContrastTheme, teamsLightTheme } from "@fluentui/react-components";
import * as microsoftTeams from "@microsoft/teams-js";
import React from "react";

export const GetFluentUITheme = () => {
  const [theme, setTheme] = React.useState(teamsLightTheme);
  microsoftTeams.getContext((context: microsoftTeams.Context) => {
    switch (context?.theme?.toLocaleLowerCase()) {
      case "light":
        setTheme(teamsLightTheme);
        break;
      case "dark":
        setTheme(teamsDarkTheme);
        break;
      case "highcontrast":
      case "contrast":
        setTheme(teamsHighContrastTheme);
        break;
    }
  });

  return theme;
};
