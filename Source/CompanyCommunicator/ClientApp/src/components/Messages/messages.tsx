// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation } from "react-i18next";
import { Spinner } from "@fluentui/react-components";
import { GetSentMessagesAction } from "../../actions";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { SentMessageDetail } from "../MessageDetail/sentMessageDetail";

export const Messages = () => {
  const { t } = useTranslation();
  const sentMessages = useAppSelector((state: RootState) => state.messages).sentMessages.payload;
  const loader = useAppSelector((state: RootState) => state.messages).isSentMessagesFetchOn.payload;
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    GetSentMessagesAction(dispatch);
  }, []);

  return (
    <>
      {loader && <Spinner labelPosition="below" label="Fetching sent messages..." />}
      {sentMessages && sentMessages.length === 0 && !loader && <div>{t("EmptySentMessages")}</div>}
      {sentMessages && sentMessages.length > 0 && !loader && <SentMessageDetail sentMessages={sentMessages} />}
    </>
  );
};
