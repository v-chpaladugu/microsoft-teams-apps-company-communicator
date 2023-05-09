// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./sendConfirmationTaskModule.scss";

import * as AdaptiveCards from "adaptivecards";
import * as React from "react";
import { useTranslation } from "react-i18next";
import { useParams } from "react-router-dom";
import { Flex } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { getConsentSummaries, getDraftNotification, sendDraftNotification } from "../../apis/messageListApi";
import { ImageUtil } from "../../utility/imageutility";
import {
  getInitAdaptiveCard,
  setCardAuthor,
  setCardBtn,
  setCardImageLink,
  setCardSummary,
  setCardTitle,
} from "../AdaptiveCard/adaptiveCard";
import { Button, MenuItem, MenuList, Text, Image, Spinner } from "@fluentui/react-components";

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
  const { id } = useParams() as any;
  let card = getInitAdaptiveCard(t);
  let adaptiveCard = new AdaptiveCards.AdaptiveCard();

  const [draftMessageItem, setDraftMessageItem] = React.useState(null);
  const [consentSummaries, setConsentSummaries] = React.useState(null);

  const [messageState, setMessageState] = React.useState<IStatusState>({
    message: { id: "", title: "" },
    loader: true,
    teamNames: [],
    rosterNames: [],
    groupNames: [],
    allUsers: false,
    messageId: 0,
  });

  React.useEffect(() => {
    getDraftMessage(id);
  }, []);

  React.useEffect(() => {
    if (draftMessageItem) {
      getConsents(id);
    }
  }, [draftMessageItem]);

  React.useEffect(() => {
    if (consentSummaries) {
      setCardTitle(card, messageState.message.title);
      setCardImageLink(card, messageState.message.imageLink);
      setCardSummary(card, messageState.message.summary);
      setCardAuthor(card, messageState.message.author);
      if (messageState.message.buttonTitle && messageState.message.buttonLink) {
        setCardBtn(card, messageState.message.buttonTitle, messageState.message.buttonLink);
      }
      adaptiveCard.parse(card);
      let renderedCard = adaptiveCard.render();
      document.getElementsByClassName("adaptiveCardContainer")[0].appendChild(renderedCard);
      if (messageState.message.buttonLink) {
        let link = messageState.message.buttonLink;
        adaptiveCard.onExecuteAction = function (action) {
          window.open(link, "_blank");
        };
      }
    }
  }, [consentSummaries]);

  const getDraftMessage = async (id: number) => {
    try {
      const response = await getDraftNotification(id);
      setDraftMessageItem(response.data);
      setMessageState({ ...messageState, message: response.data });
    } catch (error) {
      return error;
    }
  };

  const getConsents = async (id: number) => {
    try {
      const response = await getConsentSummaries(id);

      setMessageState({
        ...messageState,
        teamNames: response.data.teamNames.sort(),
        rosterNames: response.data.rosterNames.sort(),
        groupNames: response.data.groupNames.sort(),
        allUsers: response.data.allUsers,
        messageId: id,
      });

      setConsentSummaries(response.data);
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
    let resultedTeams: any[] = [];
    if (items) {
      items.map((element) => {
        resultedTeams.push(
          <MenuItem icon={<Image src={ImageUtil.makeInitialImage(element)} />}>{element}</MenuItem>
        );
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
          <MenuList>{getItemList(messageState.teamNames)}</MenuList>
        </div>
      );
    } else if (messageState.rosterNames && messageState.rosterNames.length > 0) {
      return (
        <div key="rosterNames">
          {" "}
          <span className="label">{t("TeamsMembersLabel")}</span>
          <MenuList>{getItemList(messageState.rosterNames)}</MenuList>
        </div>
      );
    } else if (messageState.groupNames && messageState.groupNames.length > 0) {
      return (
        <div key="groupNames">
          {" "}
          <span className="label">{t("GroupsMembersLabel")}</span>
          <MenuList>{getItemList(messageState.groupNames)}</MenuList>
        </div>
      );
    } else if (messageState.allUsers) {
      return (
        <div key="allUsers">
          <span className="label">{t("AllUsersLabel")}</span>
          <div className="noteText">
            <Text>{t("SendToAllUsersNote")}</Text>
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
          <Spinner />
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
                  <Spinner
                    id="sendingLoader"
                    className="hiddenLoader sendingLoader"
                    size="small"
                    label={t("PreparingMessageLabel")}
                    labelPosition="after"
                  />
                </Flex.Item>
                <Button id="sendBtn" onClick={onSendMessage} appearance="primary">{t("Send")}</Button>
              </Flex>
            </Flex>
          </Flex>
        </div>
      )}
    </>
  );
};
