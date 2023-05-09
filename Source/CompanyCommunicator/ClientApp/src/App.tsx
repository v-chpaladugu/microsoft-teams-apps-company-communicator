// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./app.scss";
import i18n from "i18next";
import React, { Suspense } from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import { FluentProvider, teamsDarkTheme, teamsHighContrastTheme, teamsLightTheme } from "@fluentui/react-components";
import * as microsoftTeams from "@microsoft/teams-js";
import Configuration from "./components/config";
import ErrorPage from "./components/ErrorPage/errorPage";
import { MainContainer } from "./components/MainContainer/mainContainer";
import { NewMessage } from "./components/NewMessage/newMessage";
import { SendConfirmationTaskModule } from "./components/SendConfirmationTaskModule/sendConfirmationTaskModule";
import SignInPage from "./components/SignInPage/signInPage";
import SignInSimpleEnd from "./components/SignInPage/signInSimpleEnd";
import SignInSimpleStart from "./components/SignInPage/signInSimpleStart";
import { StatusTaskModule } from "./components/StatusTaskModule/statusTaskModule";

export const App = () => {
  const [fluentUITheme, setFluentUITheme] = React.useState(teamsLightTheme);
  const [locale, setLocale] = React.useState("en-US");

  React.useEffect(() => {
    microsoftTeams.getContext((context: microsoftTeams.Context) => {
      let theme = context.theme || "default";
      setLocale(context.locale);
      i18n.changeLanguage(context.locale);
      updateTheme(theme);
    });
    microsoftTeams.registerOnThemeChangeHandler((theme: any) => {
      updateTheme(theme);
    });
  }, []);

  const updateTheme = (theme: string) => {
    switch (theme.toLocaleLowerCase()) {
      case "default":
        setFluentUITheme(teamsLightTheme);
        break;
      case "dark":
        setFluentUITheme(teamsDarkTheme);
        break;
      case "contrast":
        setFluentUITheme(teamsHighContrastTheme);
        break;
    }
  };

  return (
    <FluentProvider theme={fluentUITheme} dir={i18n.dir(locale)}>
      <Suspense fallback={<div></div>}>
        <div className="appContainer">
          <BrowserRouter>
            <Switch>
              <Route exact path="/configtab" component={Configuration} />
              <Route exact path="/messages" render={() => <MainContainer theme={fluentUITheme} />} />
              <Route exact path="/newmessage" component={NewMessage} />
              <Route exact path="/newmessage/:id" component={NewMessage} />
              <Route exact path="/viewstatus/:id" component={StatusTaskModule} />
              <Route exact path="/sendconfirmation/:id" component={SendConfirmationTaskModule} />
              <Route exact path="/errorpage" component={ErrorPage} />
              <Route exact path="/errorpage/:id" component={ErrorPage} />
              <Route exact path="/signin" component={SignInPage} />
              <Route exact path="/signin-simple-start" component={SignInSimpleStart} />
              <Route exact path="/signin-simple-end" component={SignInSimpleEnd} />
            </Switch>
          </BrowserRouter>
        </div>
      </Suspense>
    </FluentProvider>
  );
};
