/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */

import fetch from "node-fetch";
import * as msal from "@azure/msal-browser";

async function loadConfig() {
  const response = await fetch("/config.json");
  return response.json();
}

const config = await loadConfig();

const msalConfig = {
  auth: {
    clientId: config.clientId,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000",
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("login").onclick = login;
  document.getElementById("logout").onclick = logout;
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function login() {
  try {
    const loginRequest = {
      scopes: ["User.Read"],
    };
    const loginResponse = await msalInstance.loginPopup(loginRequest);
    console.log("Login successful", loginResponse);
    // Save tokens or user information as needed
  } catch (error) {
    console.error("Login error: ", error);
  }
}

async function logout() {
  try {
    await msalInstance.logout();
    console.log("Logout successful");
  } catch (error) {
    console.error("Logout error: ", error);
  }
}
