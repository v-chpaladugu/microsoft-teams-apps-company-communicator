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

import parse from 'html-react-parser';

export interface IMessageState {
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

  isDraftMsgUpdated: boolean;
}

export interface IConsentState {
  teamNames: string[];
  rosterNames: string[];
  groupNames: string[];
  allUsers: boolean;
  messageId: number;

  isConsentsUpdated: boolean;
}

let card: any;
let renderCard: any;

export const SendConfirmationTaskModule = () => {
  const { t } = useTranslation();
  const { id } = useParams() as any;
  const [loader, setLoader] = React.useState(true);
  const [isCardReady, setIsCardReady] = React.useState(false);

  const [messageState, setMessageState] = React.useState<IMessageState>({
    id: "",
    title: "",
    isDraftMsgUpdated: false,
  });

  const [consentState, setConsentState] = React.useState<IConsentState>({
    teamNames: [],
    rosterNames: [],
    groupNames: [],
    allUsers: false,
    messageId: 0,
    isConsentsUpdated: false,
  });

  React.useEffect(() => {
    if (id) {
      getDraftMessage(id);
      getConsents(id);
    }
  }, [id]);


  React.useEffect(() => {
    if (isCardReady && consentState.isConsentsUpdated && messageState.isDraftMsgUpdated) {
      var adaptiveCard = new AdaptiveCards.AdaptiveCard();
      adaptiveCard.parse(card);
      renderCard = adaptiveCard.render();
      if (messageState.buttonLink) {
        let link = messageState.buttonLink;
        adaptiveCard.onExecuteAction = function (action) {
          window.open(link, "_blank");
        };
      }
      setLoader(false);
    }
  }, [consentState, isCardReady, consentState.isConsentsUpdated, messageState.isDraftMsgUpdated]);


  const updateCardData = (msg: IMessageState) => {
    card = getInitAdaptiveCard(t);
    setCardTitle(card, msg.title);
    setCardImageLink(card, msg.imageLink);
    setCardSummary(card, msg.summary);
    setCardAuthor(card, msg.author);
    if (msg.buttonTitle && msg.buttonLink) {
      setCardBtn(card, msg.buttonTitle, msg.buttonLink);
    }
    setIsCardReady(true);
  };

  const getDraftMessage = async (id: number) => {
    try {
      await getDraftNotification(id).then((response) => {
        updateCardData(response.data);
        setMessageState({ ...response.data, isDraftMsgUpdated: true });
      });
    } catch (error) {
      return error;
    }
  };

  const getConsents = async (id: number) => {
    try {
      await getConsentSummaries(id).then((response) => {
        setConsentState({
          ...consentState,
          teamNames: response.data.teamNames.sort(),
          rosterNames: response.data.rosterNames.sort(),
          groupNames: response.data.groupNames.sort(),
          allUsers: response.data.allUsers,
          messageId: id,
          isConsentsUpdated: true,
        });
      });
    } catch (error) {
      return error;
    }
  };
  
  const onSendMessage = () => {
    sendDraftNotification(messageState).then(() => {
      microsoftTeams.tasks.submitTask();
    });
  };

  const getItemList = (items: string[]) => {
    let resultedTeams: any[] = [];
    if (items) {
      items.map((element) => {
        resultedTeams.push(<MenuItem icon={<Image src={ImageUtil.makeInitialImage(element)} />}>{element}</MenuItem>);
      });
    }
    return resultedTeams;
  };

  const renderAudienceSelection = () => {
    if (consentState.teamNames && consentState.teamNames.length > 0) {
      return (
        <div key="teamNames">
          {" "}
          <span className="label">{t("TeamsLabel")}</span>
          <MenuList>{getItemList(consentState.teamNames)}</MenuList>
        </div>
      );
    } else if (consentState.rosterNames && consentState.rosterNames.length > 0) {
      return (
        <div key="rosterNames">
          {" "}
          <span className="label">{t("TeamsMembersLabel")}</span>
          <MenuList>{getItemList(consentState.rosterNames)}</MenuList>
        </div>
      );
    } else if (consentState.groupNames && consentState.groupNames.length > 0) {
      return (
        <div key="groupNames">
          {" "}
          <span className="label">{t("GroupsMembersLabel")}</span>
          <MenuList>{getItemList(consentState.groupNames)}</MenuList>
        </div>
      );
    } else if (consentState.allUsers) {
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
      {loader && (
        <div className="Loader">
          <Spinner />
        </div>
      )}
      {!loader && (
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
                <div className="adaptiveCardContainer">{parse(renderCard.outerHTML)}</div>
              </Flex.Item>
            </Flex>
            <Flex className="footerContainer" vAlign="end" hAlign="end">
              <Flex className="buttonContainer" gap="gap.small">
                <Button id="sendBtn" onClick={onSendMessage} appearance="primary">
                  {t("Send")}
                </Button>
              </Flex>
            </Flex>
          </Flex>
        </div>
      )}
    </>
  );
};
