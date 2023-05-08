// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import axios from "./axiosJWTDecorator";
import { getBaseUrl } from "../configVariables";

let baseAxiosUrl = getBaseUrl() + "/api";

export const getSentNotifications = async (): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications";
  return await axios.get(url);

  // return new Promise((resolve, reject) => {
  //   resolve({
  //     data: [
  //       {
  //         id: "2517196733251094806",
  //         title: "Test3",
  //         createdDateTime: "2023-04-27T20:30:56.0123579Z",
  //         sentDate: null,
  //         succeeded: 0,
  //         failed: 0,
  //         unknown: null,
  //         canceled: null,
  //         totalMessageCount: 0,
  //         sendingStartedDate: "2023-04-27T20:31:14.8905768Z",
  //         status: "Queued",
  //         createdBy: "admin@M365x54982965.onmicrosoft.com",
  //       },
  //       {
  //         id: "2517196737161516928",
  //         title: "Test2",
  //         createdDateTime: "2023-04-27T20:24:03.8211509Z",
  //         sentDate: null,
  //         succeeded: 0,
  //         failed: 0,
  //         unknown: null,
  //         canceled: null,
  //         totalMessageCount: 0,
  //         sendingStartedDate: "2023-04-27T20:24:43.8483331Z",
  //         status: "Queued",
  //         createdBy: "admin@M365x54982965.onmicrosoft.com",
  //       },
  //       {
  //         id: "2517196738053205683",
  //         title: "Test",
  //         createdDateTime: "2023-04-27T20:21:05.889356Z",
  //         sentDate: null,
  //         succeeded: 0,
  //         failed: 0,
  //         unknown: null,
  //         canceled: null,
  //         totalMessageCount: 0,
  //         sendingStartedDate: "2023-04-27T20:23:14.6794558Z",
  //         status: "Queued",
  //         createdBy: "admin@M365x54982965.onmicrosoft.com",
  //       },
  //       {
  //         id: "2517196740739380774",
  //         title: "New test message",
  //         createdDateTime: "2023-04-27T20:18:28.0606177Z",
  //         sentDate: null,
  //         succeeded: 0,
  //         failed: 0,
  //         unknown: null,
  //         canceled: null,
  //         totalMessageCount: 0,
  //         sendingStartedDate: "2023-04-27T20:18:46.0620861Z",
  //         status: "Queued",
  //         createdBy: "admin@M365x54982965.onmicrosoft.com",
  //       },
  //       {
  //         id: "2517197457145211029",
  //         title: "testing",
  //         createdDateTime: "2023-04-27T00:24:24.1863741Z",
  //         sentDate: null,
  //         succeeded: 0,
  //         failed: 0,
  //         unknown: null,
  //         canceled: null,
  //         totalMessageCount: 0,
  //         sendingStartedDate: "2023-04-27T00:24:45.4790352Z",
  //         status: "Queued",
  //         createdBy: "admin@M365x54982965.onmicrosoft.com",
  //       },
  //       {
  //         id: "2517197815119035954",
  //         title: "This is a test message title",
  //         createdDateTime: "2023-04-26T14:27:48.6710508Z",
  //         sentDate: "2023-04-26T14:28:38.1480408Z",
  //         succeeded: 1,
  //         failed: 0,
  //         unknown: null,
  //         canceled: null,
  //         totalMessageCount: 1,
  //         sendingStartedDate: "2023-04-26T14:28:08.0966003Z",
  //         status: "Sent",
  //         createdBy: "admin@M365x54982965.onmicrosoft.com",
  //       },
  //     ],
  //   });
  // });
};

export const getDraftNotifications = async (): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications";
  return await axios.get(url);

  // return new Promise((resolve, reject) => {
  //   resolve({ data: [{ id: "1234", title: "test" }] });
  // });
};

export const verifyGroupAccess = async (): Promise<any> => {
  let url = baseAxiosUrl + "/groupdata/verifyaccess";
  return await axios.get(url, false);
};

export const getGroups = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/groupdata/" + id;
  return await axios.get(url);
};

export const searchGroups = async (query: string): Promise<any> => {
  let url = baseAxiosUrl + "/groupdata/search/" + query;
  return await axios.get(url);
};

export const exportNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/exportnotification/export";
  return await axios.put(url, payload);
};

export const getSentNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications/" + id;
  return await axios.get(url);
};

export const getDraftNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/" + id;
  return await axios.get(url);
};

export const deleteDraftNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/" + id;
  return await axios.delete(url);
};

export const duplicateDraftNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/duplicates/" + id;
  return await axios.post(url);
};

export const sendDraftNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications";
  return await axios.post(url, payload);
};

export const updateDraftNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications";
  return await axios.put(url, payload);
};

export const createDraftNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications";
  return await axios.post(url, payload);
};

export const getTeams = async (): Promise<any> => {
  let url = baseAxiosUrl + "/teamdata";
  return await axios.get(url);
};

export const cancelSentNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications/cancel/" + id;
  return await axios.post(url);
};

export const getConsentSummaries = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/consentSummaries/" + id;
  return await axios.get(url);
};

export const sendPreview = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/draftnotifications/previews";
  return await axios.post(url, payload);
};

export const getAuthenticationConsentMetadata = async (
  windowLocationOriginDomain: string,
  login_hint: string
): Promise<any> => {
  let url = `${baseAxiosUrl}/authenticationMetadata/consentUrl?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${login_hint}`;
  return await axios.get(url, undefined, false);
};
