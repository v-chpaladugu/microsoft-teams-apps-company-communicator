import { createSlice } from '@reduxjs/toolkit'

export interface MessagesState {
    draftMessages: any;
    sentMessages: any;
    selectedMessage: any;
  }

const initialState: MessagesState = {
    draftMessages: [], sentMessages: [], selectedMessage: {}
}

export const messagesSlice = createSlice({
  name: 'messagesSlice',
  initialState,
  reducers: {
    draftMessagesReducer: (state, action) => {
      state.draftMessages = action.payload
    },
    sentMessagesReducer: (state, action) => {
        state.sentMessages = action.payload
    },
    selectedMessageReducer: (state, action) => {
        state.selectedMessage = action.payload
    },
  },
})

export const { draftMessagesReducer, sentMessagesReducer, selectedMessageReducer } = messagesSlice.actions

export default messagesSlice.reducer
