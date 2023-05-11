import { createSlice } from "@reduxjs/toolkit";

export interface MessagesState {
  draftMessages: { action: string; payload: [] };
  sentMessages: { action: string; payload: [] };
  selectedMessage: { action: string; payload: {} };
  teamsData: { action: string; payload: [] };
  groups: { action: string; payload: [] },
  verifyGroup: { action: string; payload: boolean },
  isDraftMessagesFetchOn: { action: string; payload: boolean },
  isSentMessagesFetchOn: { action: string; payload: boolean }
}

const initialState: MessagesState = {
  draftMessages: { action: "FETCH_DRAFT_MESSAGES", payload: [] },
  sentMessages: { action: "FETCH_MESSAGES", payload: [] },
  selectedMessage: { action: "MESSAGE_SELECTED", payload: [] },
  teamsData: { action: "GET_TEAMS_DATA", payload: [] },
  groups: { action: "GET_GROUPS", payload: [] },
  verifyGroup: { action: "VERIFY_GROUP_ACCESS", payload: false },
  isDraftMessagesFetchOn: { action: "DRAFT_MESSAGES_FETCH_STATUS", payload: false },
  isSentMessagesFetchOn: { action: "SENT_MESSAGES_FETCH_STATUS", payload: false }
};

export const messagesSlice = createSlice({
  name: "messagesSlice",
  initialState,
  reducers: {
    draftMessages: (state, action) => {
      state.draftMessages = action.payload;
    },
    sentMessages: (state, action) => {
      state.sentMessages = action.payload;
    },
    selectedMessage: (state, action) => {
      state.selectedMessage = action.payload;
    },
    teamsData: (state, action) => {
      state.teamsData = action.payload;
    },
    groups: (state, action) => {
      state.groups = action.payload;
    },
    verifyGroup: (state, action) => {
      state.groups = action.payload;
    },
    isDraftMessagesFetchOn: (state, action) => {
      state.isDraftMessagesFetchOn = action.payload;
    },
    isSentMessagesFetchOn: (state, action) => {
      state.isSentMessagesFetchOn = action.payload;
    },
  },
});

export const { draftMessages, sentMessages, selectedMessage, teamsData, groups, verifyGroup, isDraftMessagesFetchOn, isSentMessagesFetchOn} = messagesSlice.actions;

export default messagesSlice.reducer;
