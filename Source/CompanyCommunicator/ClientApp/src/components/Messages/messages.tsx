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
  const [loader, setLoader] = React.useState(true);
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    GetSentMessagesAction(dispatch);
  }, []);

  React.useEffect(() => {
    if (sentMessages && sentMessages.length > 0) {
      setLoader(false);
    }
  }, [sentMessages]);

  return (
    <>
      {loader && <Spinner labelPosition="below" size="large" label="Fetching sent messages..." />}
      {sentMessages && sentMessages.length === 0 && !loader && <div>{t("EmptySentMessages")}</div>}
      {sentMessages && sentMessages.length > 0 && !loader && <SentMessageDetail sentMessages={sentMessages} />}
    </>
  );
};
