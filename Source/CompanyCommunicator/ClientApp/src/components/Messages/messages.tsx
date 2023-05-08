// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { connect } from "react-redux";
import { useTranslation, withTranslation, WithTranslation } from "react-i18next";
import { TooltipHost } from "office-ui-fabric-react";
import {
  Loader,
  List,
  Flex,
  Text,
  AcceptIcon,
  CloseIcon,
  ExclamationCircleIcon,
  ExclamationTriangleIcon,
} from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";

import { SelectedMessageAction, GetSentMessagesAction, GetDraftMessagesAction } from "../../actions";
import { getBaseUrl } from "../../configVariables";
import Overflow from "../OverFlow/sentMessageOverflow";
import "./messages.scss";
import { TFunction } from "i18next";
import { formatNumber } from "../../i18n";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { Spinner } from "@fluentui/react-components";
import { SentMessageDetail } from "../OverFlow/sentMessageDetail";

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
  title: string;
  sentDate: string;
  recipients: string;
  acknowledgements?: string;
  reactions?: string;
  responses?: string;
}

export interface IMessageProps extends WithTranslation {
  messagesList: IMessage[];
  selectMessage?: any;
  getMessagesList?: any;
  getDraftMessagesList?: any;
}

export interface IMessageState {
  message: IMessage[];
  loader: boolean;
}

export const Messages = () => {
  const { t } = useTranslation();
  // private interval: any;
  const [isOpenTaskModuleAllowed, setIsOpenTaskModuleAllowed] = React.useState(true);
  const sentMessages = useAppSelector((state: RootState) => state.messages).sentMessages.payload;
  const [loader, setLoader] = React.useState(true);
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    //- Handle the Esc key
    document.addEventListener("keydown", escFunction, false);
    GetSentMessagesAction(dispatch);
    // // tslint:disable-next-line no-string-based-set-interval
    // this.interval = setInterval(() => {
    //     this.props.getMessagesList();
    // }, 60000);
  }, []);

  React.useEffect(() => {
    if (sentMessages && sentMessages.length > 0) {
      setLoader(false);
    }
  }, [sentMessages]);

  //   let keyCount = 0;
  //   const processItem = (message: any) => {
  //       keyCount++;
  //       const out = {
  //           key: keyCount,
  //           content: this.messageContent(message),
  //           onClick: (): void => {
  //               let url = getBaseUrl() + "/viewstatus/" + message.id + "?locale={locale}";
  //               this.onOpenTaskModule(null, url, this.localize("ViewStatus"));
  //           },
  //           styles: { margin: '0.2rem 0.2rem 0 0' },
  //       };
  //       return out;
  //   };

  //   const label = this.processLabels();
  //   const outList = this.state.message.map(processItem);
  //   const allMessages = [...label, ...outList];

  const escFunction = (event: any) => {
    if (event.keyCode === 27 || event.key === "Escape") {
      microsoftTeams.tasks.submitTask();
    }
  };

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
      };

      microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }
  };

  return (
    <>
      {loader && <Spinner labelPosition="below" size="large" label="Fetching sent messages..." />}
      {sentMessages && sentMessages.length === 0 && !loader && <div className="results">{t("EmptySentMessages")}</div>}
      {sentMessages && sentMessages.length > 0 && !loader && <SentMessageDetail sentMessages={sentMessages} />}
    </>
  );

  // private processLabels = () => {
  //     const out = [{
  //         key: "labels",
  //         content: (
  //             <Flex vAlign="center" fill gap="gap.small">
  //                 <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} grow={1} >
  //                     <Text
  //                         truncated
  //                         weight="bold"
  //                         content={this.localize("TitleText")}
  //                     >
  //                     </Text>
  //                 </Flex.Item>
  //                 <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }}>
  //                     <Text></Text>
  //                 </Flex.Item>
  //                 <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} shrink={false}>
  //                     <Text
  //                         truncated
  //                         content={this.localize("Recipients")}
  //                         weight="bold"
  //                     >
  //                     </Text>
  //                 </Flex.Item>
  //                 <Flex.Item size="size.quarter" variables={{ 'size.quarter': '16%' }} >
  //                     <Text
  //                         truncated
  //                         content={this.localize("Sent")}
  //                         weight="bold"
  //                     >
  //                     </Text>
  //                 </Flex.Item>
  //                 <Flex.Item size="size.quarter" variables={{ 'size.quarter': '16%' }} >
  //                     <Text
  //                         truncated
  //                         content={this.localize("CreatedBy")}
  //                         weight="bold"
  //                     >
  //                     </Text>
  //                 </Flex.Item>
  //                 <Flex.Item shrink={0} >
  //                     <Overflow title="" />
  //                 </Flex.Item>
  //             </Flex>
  //         ),
  //         styles: { margin: '0.2rem 0.2rem 0 0' },
  //     }];
  //     return out;
  // }

  // private renderSendingText = (message: any) => {
  //     var text = "";
  //     switch (message.status) {
  //         case "Queued":
  //             text = this.localize("Queued");
  //             break;
  //         case "SyncingRecipients":
  //             text = this.localize("SyncingRecipients");
  //             break;
  //         case "InstallingApp":
  //             text = this.localize("InstallingApp");
  //             break;
  //         case "Sending":
  //             let sentCount =
  //                 (message.succeeded ? message.succeeded : 0) +
  //                 (message.failed ? message.failed : 0) +
  //                 (message.unknown ? message.unknown : 0);

  //             text = this.localize("SendingMessages", { "SentCount": formatNumber(sentCount), "TotalCount": formatNumber(message.totalMessageCount) });
  //             break;
  //         case "Canceling":
  //             text = this.localize("Canceling");
  //             break;
  //         case "Canceled":
  //         case "Sent":
  //         case "Failed":
  //             text = "";
  //     }

  //     return (<Text truncated content={text} />);
  // }

  // private messageContent = (message: any) => {
  //     return (
  //         <Flex className="listContainer" vAlign="center" fill gap="gap.small">
  //             <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} grow={1}>
  //                 <Text
  //                     truncated
  //                     content={message.title}
  //                 >
  //                 </Text>
  //             </Flex.Item>
  //             <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }}>
  //                 {this.renderSendingText(message)}
  //             </Flex.Item>
  //             <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} shrink={false}>
  //                 <div>
  //                     <TooltipHost content={this.props.t("TooltipSuccess")} calloutProps={{ gapSpace: 0 }}>
  //                         <AcceptIcon xSpacing="after" className="succeeded" outline />
  //                         <span className="semiBold">{formatNumber(message.succeeded)}</span>
  //                     </TooltipHost>
  //                     <TooltipHost content={this.props.t("TooltipFailure")} calloutProps={{ gapSpace: 0 }}>
  //                         <CloseIcon xSpacing="both" className="failed" outline />
  //                         <span className="semiBold">{formatNumber(message.failed)}</span>
  //                     </TooltipHost>
  //                     {
  //                         message.canceled &&
  //                         <TooltipHost content="Canceled" calloutProps={{ gapSpace: 0 }}>
  //                             <ExclamationCircleIcon xSpacing="both" className="canceled" outline />
  //                             <span className="semiBold">{formatNumber(message.canceled)}</span>
  //                         </TooltipHost>
  //                     }
  //                     {
  //                         message.unknown &&
  //                         <TooltipHost content="Unknown" calloutProps={{ gapSpace: 0 }}>
  //                             <ExclamationTriangleIcon xSpacing="both" className="unknown" outline />
  //                             <span className="semiBold">{formatNumber(message.unknown)}</span>
  //                         </TooltipHost>
  //                     }
  //                 </div>
  //             </Flex.Item>
  //             <Flex.Item size="size.quarter" variables={{ 'size.quarter': '16%' }} >
  //                 <Text
  //                     truncated
  //                     className="semiBold"
  //                     content={message.sentDate}
  //                 />
  //             </Flex.Item>
  //             <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} >
  //                 <Text
  //                     truncated
  //                     className="semiBold"
  //                     content={message.createdBy}
  //                 />
  //             </Flex.Item>
  //             <Flex.Item shrink={0}>
  //                 <Overflow message={message} title="" />
  //             </Flex.Item>
  //         </Flex>
  //     );
  // }
};
