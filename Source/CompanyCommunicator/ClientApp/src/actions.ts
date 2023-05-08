// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { getDraftNotifications, getSentNotifications } from "./apis/messageListApi";
import { formatDate } from "./i18n";
import { draftMessages, selectedMessage, sentMessages } from "./messagesSlice";
import { store } from "./store";

type Notification = {
  createdDateTime: string;
  failed: number;
  id: string;
  isCompleted: boolean;
  sentDate: string;
  sendingStartedDate: string;
  sendingDuration: string;
  succeeded: number;
  throttled: number;
  title: string;
  totalMessageCount: number;
  createdBy: string;
};

export const SelectedMessageAction = (dispatch: typeof store.dispatch, message: any) => {
  dispatch(selectedMessage({ type: "MESSAGE_SELECTED", payload: message }));
};

export const GetSentMessagesAction = (dispatch: typeof store.dispatch) => {
  getSentNotifications().then((response) => {
    const notificationList: Notification[] = response.data;
    notificationList.forEach((notification) => {
      notification.sendingStartedDate = formatDate(notification.sendingStartedDate);
      notification.sentDate = formatDate(notification.sentDate);
    });
    dispatch(sentMessages({ type: "FETCH_MESSAGES", payload: notificationList }));
  });
};

export const GetDraftMessagesAction = (dispatch: typeof store.dispatch) => {
  getDraftNotifications().then((response) => {
    dispatch(draftMessages({ type: "FETCH_DRAFT_MESSAGES", payload: response.data }));
  });
};
