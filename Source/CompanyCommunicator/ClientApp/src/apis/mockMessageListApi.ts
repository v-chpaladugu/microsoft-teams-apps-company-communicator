// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import axios from "./axiosJWTDecorator";
import { getBaseUrl } from "../configVariables";

let baseAxiosUrl = getBaseUrl() + "/api";

export const getSentNotifications = async (): Promise<any> => {
  // let url = baseAxiosUrl + "/sentnotifications";
  // return await axios.get(url);

  return new Promise((resolve, reject) => {
    resolve({
      data: [
        {
          id: "2517196733251094806",
          title: "Test3",
          createdDateTime: "2023-04-27T20:30:56.0123579Z",
          sentDate: null,
          succeeded: 0,
          failed: 0,
          unknown: null,
          canceled: null,
          totalMessageCount: 0,
          sendingStartedDate: "2023-04-27T20:31:14.8905768Z",
          status: "Queued",
          createdBy: "admin@M365x54982965.onmicrosoft.com",
        },
        {
          id: "2517196737161516928",
          title: "Test2",
          createdDateTime: "2023-04-27T20:24:03.8211509Z",
          sentDate: null,
          succeeded: 0,
          failed: 0,
          unknown: null,
          canceled: null,
          totalMessageCount: 0,
          sendingStartedDate: "2023-04-27T20:24:43.8483331Z",
          status: "Queued",
          createdBy: "admin@M365x54982965.onmicrosoft.com",
        },
        {
          id: "2517196738053205683",
          title: "Test",
          createdDateTime: "2023-04-27T20:21:05.889356Z",
          sentDate: null,
          succeeded: 0,
          failed: 0,
          unknown: null,
          canceled: null,
          totalMessageCount: 0,
          sendingStartedDate: "2023-04-27T20:23:14.6794558Z",
          status: "Queued",
          createdBy: "admin@M365x54982965.onmicrosoft.com",
        },
        {
          id: "2517196740739380774",
          title: "New test message",
          createdDateTime: "2023-04-27T20:18:28.0606177Z",
          sentDate: null,
          succeeded: 0,
          failed: 0,
          unknown: null,
          canceled: null,
          totalMessageCount: 0,
          sendingStartedDate: "2023-04-27T20:18:46.0620861Z",
          status: "Queued",
          createdBy: "admin@M365x54982965.onmicrosoft.com",
        },
        {
          id: "2517197457145211029",
          title: "testing",
          createdDateTime: "2023-04-27T00:24:24.1863741Z",
          sentDate: null,
          succeeded: 0,
          failed: 0,
          unknown: null,
          canceled: null,
          totalMessageCount: 0,
          sendingStartedDate: "2023-04-27T00:24:45.4790352Z",
          status: "Queued",
          createdBy: "admin@M365x54982965.onmicrosoft.com",
        },
        {
          id: "2517197815119035954",
          title: "This is a test message title",
          createdDateTime: "2023-04-26T14:27:48.6710508Z",
          sentDate: "2023-04-26T14:28:38.1480408Z",
          succeeded: 1,
          failed: 0,
          unknown: null,
          canceled: null,
          totalMessageCount: 1,
          sendingStartedDate: "2023-04-26T14:28:08.0966003Z",
          status: "Sent",
          createdBy: "admin@M365x54982965.onmicrosoft.com",
        },
      ],
    });
  });
};

export const getDraftNotifications = async (): Promise<any> => {
  // let url = baseAxiosUrl + "/draftnotifications";
  // return await axios.get(url);

  return new Promise((resolve, reject) => {
    resolve({ data: [{"id":"0638187395972683184","title":"Test 123"},{"id":"0638192293612328690","title":"Test 123 (copy)"}] });
  });
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
  // let url = baseAxiosUrl + "/groupdata/search/" + query;
  // return await axios.get(url);

  var test = [{"id":"19:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"1-Test-cc-raj"}, 
  {"id":"20:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"AATcc-raj"}, 
  {"id":"21:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC-raj"},
  {"id":"22:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC-cc-raj"},
  {"id":"23:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"Raj-cc-raj"},
  {"id":"24:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"XYZ-Team"},

  {"id":"25:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"BB1-Test-cc-raj"}, 
  {"id":"26:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"BATcc-raj"}, 
  {"id":"27:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC-raj"},
  {"id":"28:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC-cc-raj"},
  {"id":"29:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"Raj-cc-raj"},
  {"id":"30:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"XYZ-Team"},

  {"id":"31:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"CA1-Test-cc-raj"}, 
  {"id":"201:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"Tcc-raj"}, 
  {"id":"212:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC-raj"},
  {"id":"223:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"XYZABC-cc-raj"},
  {"id":"234:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"Raj-cc-raj"},
  {"id":"244:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"XYZ-Team"},

  {"id":"519:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC1-Test-cc-raj"}, 
  {"id":"520:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"TETcc-raj"}, 
  {"id":"521:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"RRRABC-raj"},
  {"id":"622:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"RABC-cc-raj"},
  {"id":"723:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ARaj-cc-raj"},
  {"id":"724:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"MRRAXYZ-Team"}];

  var result = test.filter(x => x.name.toLowerCase().includes(query.toLowerCase()));

  return new Promise((resolve, reject) => {
    resolve({ data: result || []});
});


};

export const exportNotification = async (payload: {}): Promise<any> => {
  let url = baseAxiosUrl + "/exportnotification/export";
  return await axios.put(url, payload);
};

export const getSentNotification = async (id: number): Promise<any> => {
  // let url = baseAxiosUrl + "/sentnotifications/" + id;
  // return await axios.get(url);

  return new Promise((resolve, reject) => {

 resolve({ data: {"sendingStartedDate":"2023-04-26T10:09:23.1439055Z","sentDate":"2023-04-26T10:09:46.2170471Z","succeeded":2,"failed":0,"unknown":null,"canceled":null,"teamNames":[],"rosterNames":[],"groupNames":[],"allUsers":true,"errorMessage":null,"warningMessage":null,"canDownload":true,"sendingCompleted":true,"createdBy":"admin@M365x75769129.onmicrosoft.com","id":"0638192293612328690","title":"Check for Test Message Option 3","imageLink":"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAMAAAAJbSJIAAAAclBMVEXz8/PzUyWBvAYFpvD/ugj19Pbz+fr39fr69vPy9fr29PPzRAB5uAAAofD/tgDz29bh6tTzTBbzmoiw0oGBxfH70IHU5vP16tTz5OHo7eDzPADzlIGs0Hnf6/N5wvH7znn07eAAnvDzvrTL3rCv1/L43rD2QPCNAAABfklEQVR4nO3cOXLCQBRFUXkQg0DMoxAIPOx/i06gCdxVjj4EPncDr0510tEvCkmS9JzK8HKr1TS6XQLuZ9HlhNNmFFtz6F2n+u04ulnmFaejl9jmiyRcvgZHSEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhI+FhhM49tdBe242cId4dFcB83YfnZLmNr97k7u73w0lTZjy57SFiSpD+rBtFVaascRpcDHk+r2E5fN2L53a1j684Z4eZSB3caXLeG3SS67e9PTbWp32KrV0m4nrzHRkhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISPhYYfT90suT75cWxSa6Yzqze95GV2WARRXefauMLgeUJOn/9gOEMUYmQwAZiQAAAABJRU5ErkJggg==", "imageBase64BlobName":"0638192293612328690","summary":"Test Summary for Test Message checking messaging options","author":"Jayant","buttonTitle":"","buttonLink":"","createdDateTime":"2023-04-26T10:09:13.1765093Z"} });

});
};

export const getDraftNotification = async (id: number): Promise<any> => {
//   let url = baseAxiosUrl + "/draftnotifications/" + id;
//   return await axios.get(url);
  
return new Promise((resolve, reject) => {
    resolve({ data: {
      "teams":["19:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2"],
      "rosters":[],
      "groups":[],
      "allUsers":false,
      "id":"0638192293612328690",
      "title":"Test 123 (copy)",
      "imageLink":"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAMAAAAJbSJIAAAAclBMVEXz8/PzUyWBvAYFpvD/ugj19Pbz+fr39fr69vPy9fr29PPzRAB5uAAAofD/tgDz29bh6tTzTBbzmoiw0oGBxfH70IHU5vP16tTz5OHo7eDzPADzlIGs0Hnf6/N5wvH7znn07eAAnvDzvrTL3rCv1/L43rD2QPCNAAABfklEQVR4nO3cOXLCQBRFUXkQg0DMoxAIPOx/i06gCdxVjj4EPncDr0510tEvCkmS9JzK8HKr1TS6XQLuZ9HlhNNmFFtz6F2n+u04ulnmFaejl9jmiyRcvgZHSEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhI+FhhM49tdBe242cId4dFcB83YfnZLmNr97k7u73w0lTZjy57SFiSpD+rBtFVaascRpcDHk+r2E5fN2L53a1j684Z4eZSB3caXLeG3SS67e9PTbWp32KrV0m4nrzHRkhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISEhISPhYYfT90suT75cWxSa6Yzqze95GV2WARRXefauMLgeUJOn/9gOEMUYmQwAZiQAAAABJRU5ErkJggg==",
      "imageBase64BlobName":"0638192293612328690",
      "summary":"Test",
      "author":"Test",
      "buttonTitle":"Test",
      "buttonLink":"https://google.com",
      "createdDateTime":"2023-05-09T11:42:41.232874Z"} });
});
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
  // let url = baseAxiosUrl + "/teamdata";
  // return await axios.get(url);

  return new Promise((resolve, reject) => {
    resolve({ data: [{"id":"19:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"1-Test-cc-raj"}, 
    {"id":"20:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"Tcc-raj"}, 
    {"id":"21:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC-raj"},
    {"id":"22:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"ABC-cc-raj"},
    {"id":"23:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"Raj-cc-raj"},
    {"id":"24:-NerdIDjIGqfzXbVO7NcJwX6MNj8irw2OhCbsfcYtoQ1@thread.tacv2","name":"XYZ-Team"}] });
});


  
};

export const cancelSentNotification = async (id: number): Promise<any> => {
  let url = baseAxiosUrl + "/sentnotifications/cancel/" + id;
  return await axios.post(url);
};

export const getConsentSummaries = async (id: number): Promise<any> => {
//   let url = baseAxiosUrl + "/draftnotifications/consentSummaries/" + id;
//   return await axios.get(url);

return new Promise((resolve, reject) => {
    resolve({ data: {"notificationId":"0638192293612328690","teamNames":["Test-cc-raj"],"rosterNames":[],"groupNames":[],"allUsers":false}});
});

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
