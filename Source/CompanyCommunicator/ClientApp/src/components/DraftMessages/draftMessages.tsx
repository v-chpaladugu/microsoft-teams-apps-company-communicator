// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { useTranslation, WithTranslation } from "react-i18next";
import { List } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";

import "./draftMessages.scss";
import { GetDraftMessagesAction, GetSentMessagesAction } from "../../actions";
import { getBaseUrl } from "../../configVariables";
// import Overflow from "../OverFlow/draftMessageOverflow";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { Spinner } from "@fluentui/react-components";
import { DraftMessageDetail } from "../OverFlow/draftMessageDetail";

export interface ITaskInfo {
  title?: string;
  height?: number;
  width?: number;
  url?: string;
  card?: string;
  fallbackUrl?: string;
  completionBotId?: string;
}

export interface IMessage {
  id: string;
  title: string;
  date: string;
  recipients: string;
  acknowledgements?: string;
  reactions?: string;
  responses?: string;
}

export interface IMessageProps extends WithTranslation {
  messages: IMessage[];
  selectedMessage: any;
  selectMessage?: any;
  getDraftMessagesList?: any;
  getMessagesList?: any;
}

export interface IMessageState {
  message: IMessage[];
  itemsAccount: number;
  loader: boolean;
  teamsTeamId?: string;
  teamsChannelId?: string;
}

export const DraftMessages = () => {
  const { t } = useTranslation();
  // let interval: any;
  const [isOpenTaskModuleAllowed, setIsOpenTaskModuleAllowed] = React.useState(true);
  const draftMessages = useAppSelector((state: RootState) => state.messages).draftMessages.payload;
  const [loader, setLoader] = React.useState(true);
  const [teamsTeamId, setTeamsTeamId] = React.useState("");
  const [teamsChannelId, setTeamsChannelId] = React.useState("");
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    microsoftTeams.getContext((context: any) => {
      setTeamsTeamId(context.teamId);
      setTeamsChannelId(context.channelId);
    });
    GetDraftMessagesAction(dispatch);
  }, []);

  React.useEffect(() => {
    if (draftMessages && draftMessages.length > 0) {
      setLoader(false);
    }
  }, [draftMessages]);

  const onOpenTaskModule = (event: any, url: string, title: string) => {
    if (isOpenTaskModuleAllowed) {
      setIsOpenTaskModuleAllowed(false);
      let taskInfo: ITaskInfo = {
        url: url,
        title: title,
        height: 530,
        width: 1000,
        fallbackUrl: url,
      };

      let submitHandler = (err: any, result: any) => {
        setIsOpenTaskModuleAllowed(true);
        GetDraftMessagesAction(dispatch);
        GetSentMessagesAction(dispatch);
      };

      microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }
  };

  //   const label = t("TitleText");
  //   const outList = draftMessages.map(processItem);
  //   const allDraftMessages = [...label, ...outList];

  //   const processItem = (message: any) => {
  //     keyCount++;
  //     const out = {
  //       key: keyCount,
  //       content: (
  //         <Flex vAlign="center" fill gap="gap.small">
  //           <Flex.Item shrink={0} grow={1}>
  //             <Text>{message.title}</Text>
  //           </Flex.Item>
  //           <Flex.Item shrink={0} align="end">
  //             <Overflow message={message} title="" />
  //           </Flex.Item>
  //         </Flex>
  //       ),
  //       styles: { margin: "0.2rem 0.2rem 0 0" },
  //       onClick: (): void => {
  //         let url = getBaseUrl() + "/newmessage/" + message.id + "?locale={locale}";
  //         this.onOpenTaskModule(null, url, this.localize("EditMessage"));
  //       },
  //     };
  //     return out;
  //   };

  return (
    <>
      {loader && <Spinner labelPosition="below" size="large" label="Fetching draft messages..." />}
      {draftMessages && draftMessages.length === 0 && !loader && <div className="results">{t("EmptyDraftMessages")}</div>}
      {draftMessages && draftMessages.length > 0 && !loader && <DraftMessageDetail draftMessages={draftMessages} />}
    </>
  );
};
