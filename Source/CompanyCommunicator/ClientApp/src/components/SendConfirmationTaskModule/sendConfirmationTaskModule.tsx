// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import { useTranslation, WithTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import { Loader, Button, Text, List, Image, Flex } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";

import "./sendConfirmationTaskModule.scss";
import { getDraftNotification, getConsentSummaries, sendDraftNotification } from "../../apis/messageListApi";
import {
  getInitAdaptiveCard,
  setCardTitle,
  setCardImageLink,
  setCardSummary,
  setCardAuthor,
  setCardBtn,
} from "../AdaptiveCard/adaptiveCard";
import { ImageUtil } from "../../utility/imageutility";
import { useParams } from "react-router-dom";

export interface IListItem {
  header: string;
  media: JSX.Element;
}

export interface IMessage {
  id: string;
  title: string;
  acknowledgements?: number;
  reactions?: number;
  responses?: number;
  succeeded?: number;
  failed?: number;
  throttled?: number;
  sentDate?: string;
  imageLink?: string;
  summary?: string;
  author?: string;
  buttonLink?: string;
  buttonTitle?: string;
  createdBy?: string;
}

export interface SendConfirmationTaskModuleProps extends RouteComponentProps, WithTranslation {}

export interface IStatusState {
  message: IMessage;
  loader: boolean;
  teamNames: string[];
  rosterNames: string[];
  groupNames: string[];
  allUsers: boolean;
  messageId: number;
}

export const SendConfirmationTaskModule = () => {
  const { t } = useTranslation();
  const [messageState, setMessageState] = React.useState<IStatusState>({
    message: { id: "", title: "" },
    loader: true,
    teamNames: [],
    rosterNames: [],
    groupNames: [],
    allUsers: false,
    messageId: 0,
  });

  const { id } = useParams() as any;

  let card: any;

  React.useEffect(() => {
    card = getInitAdaptiveCard(t);
    initialLoad();
  }, []);

  const initialLoad = () => {
    if (id) {
      getItem(id).then(() => {
        getConsentSummaries(id).then((response) => {
          setMessageState({
            ...messageState,
            teamNames: response.data.teamNames.sort(),
            rosterNames: response.data.rosterNames.sort(),
            groupNames: response.data.groupNames.sort(),
            allUsers: response.data.allUsers,
            messageId: id,
          });

          setMessageState({ ...messageState, loader: false });

          setCardTitle(card, messageState.message.title);
          setCardImageLink(card, messageState.message.imageLink);
          setCardSummary(card, messageState.message.summary);
          setCardAuthor(card, messageState.message.author);
          if (messageState.message.buttonTitle && messageState.message.buttonLink) {
            setCardBtn(card, messageState.message.buttonTitle, messageState.message.buttonLink);
          }

          let adaptiveCard = new AdaptiveCards.AdaptiveCard();
          adaptiveCard.parse(card);
          let renderedCard = adaptiveCard.render();
          document.getElementsByClassName("adaptiveCardContainer")[0].appendChild(renderedCard);
          if (messageState.message.buttonLink) {
            let link = messageState.message.buttonLink;
            adaptiveCard.onExecuteAction = function (action) {
              window.open(link, "_blank");
            };
          }
        });
      });
    }
  };

  const getItem = async (id: number) => {
    try {
      const response = await getDraftNotification(id);
      setMessageState({ ...messageState, message: response.data });
    } catch (error) {
      return error;
    }
  };

  const onSendMessage = () => {
    let spanner = document.getElementsByClassName("sendingLoader");
    spanner[0].classList.remove("hiddenLoader");
    sendDraftNotification(messageState.message).then(() => {
      microsoftTeams.tasks.submitTask();
    });
  };

  const getItemList = (items: string[]) => {
    let resultedTeams: IListItem[] = [];
    if (items) {
      resultedTeams = items.map((element) => {
        const resultedTeam: IListItem = {
          header: element,
          media: <Image src={ImageUtil.makeInitialImage(element)} avatar />,
        };
        return resultedTeam;
      });
    }
    return resultedTeams;
  };

  const renderAudienceSelection = () => {
    if (messageState.teamNames && messageState.teamNames.length > 0) {
      return (
        <div key="teamNames">
          {" "}
          <span className="label">{t("TeamsLabel")}</span>
          <List items={getItemList(messageState.teamNames)} />
        </div>
      );
    } else if (messageState.rosterNames && messageState.rosterNames.length > 0) {
      return (
        <div key="rosterNames">
          {" "}
          <span className="label">{t("TeamsMembersLabel")}</span>
          <List items={getItemList(messageState.rosterNames)} />
        </div>
      );
    } else if (messageState.groupNames && messageState.groupNames.length > 0) {
      return (
        <div key="groupNames">
          {" "}
          <span className="label">{t("GroupsMembersLabel")}</span>
          <List items={getItemList(messageState.groupNames)} />
        </div>
      );
    } else if (messageState.allUsers) {
      return (
        <div key="allUsers">
          <span className="label">{t("AllUsersLabel")}</span>
          <div className="noteText">
            <Text error content={t("SendToAllUsersNote")} />
          </div>
        </div>
      );
    } else {
      return <div></div>;
    }
  };

  return (
    <>
      {messageState.loader && (
        <div className="Loader">
          <Loader />
        </div>
      )}
      {!messageState.loader && (
        <div className="taskModule">
          <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
            <Flex className="scrollableContent" gap="gap.small">
              <Flex.Item size="size.half">
                <Flex column className="formContentContainer">
                  <h3>{t("ConfirmToSend")}</h3>
                  <span>{t("SendToRecipientsLabel")}</span>
                  <div className="results">{renderAudienceSelection()}</div>
                </Flex>
              </Flex.Item>
              <Flex.Item size="size.half">
                <div className="adaptiveCardContainer"></div>
              </Flex.Item>
            </Flex>
            <Flex className="footerContainer" vAlign="end" hAlign="end">
              <Flex className="buttonContainer" gap="gap.small">
                <Flex.Item push>
                  <Loader
                    id="sendingLoader"
                    className="hiddenLoader sendingLoader"
                    size="smallest"
                    label={t("PreparingMessageLabel")}
                    labelPosition="end"
                  />
                </Flex.Item>
                <Button content={t("Send")} id="sendBtn" onClick={onSendMessage} primary />
              </Flex>
            </Flex>
          </Flex>
        </div>
      )}
    </>
  );
};
