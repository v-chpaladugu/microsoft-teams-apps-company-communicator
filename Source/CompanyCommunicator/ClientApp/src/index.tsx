// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from "react";
import ReactDOM from "react-dom";
import { HelmetProvider } from "react-helmet-async";
import { Provider } from "react-redux";
import * as microsoftTeams from "@microsoft/teams-js";
import { App } from "./App";
import * as serviceWorker from "./serviceWorker";
import { store } from "./store";

microsoftTeams.initialize();

const helmetContext = {};

ReactDOM.render(
  <HelmetProvider context={helmetContext}>
    <Provider store={store}>
      <App />
    </Provider>
  </HelmetProvider>,
  document.getElementById("root")
);

// If you want your app to work offline and load faster, you can change
// unregister() to register() below. Note this comes with some pitfalls.
// Learn more about service workers: https://bit.ly/CRA-PWA
serviceWorker.unregister();
