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
    draftMessages: (state, action) => {
      state.draftMessages = action.payload
    },
    sentMessages: (state, action) => {
        state.sentMessages = action.payload
    },
    selectedMessage: (state, action) => {
        state.selectedMessage = action.payload
    },
  },
})

export const { draftMessages, sentMessages, selectedMessage } = messagesSlice.actions

export default messagesSlice.reducer

// The function below is called a thunk and allows us to perform async logic. It
// can be dispatched like a regular action: `dispatch(incrementAsync(10))`. This
// will call the thunk with the `dispatch` function as the first argument. Async
// code can then be executed and other actions can be dispatched
// export const incrementAsync = (amount) => (dispatch) => {
//   setTimeout(() => {
//     dispatch(incrementByAmount(amount))
//   }, 1000)
// }