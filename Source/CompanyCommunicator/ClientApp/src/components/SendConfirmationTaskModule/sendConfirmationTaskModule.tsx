// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as AdaptiveCards from "adaptivecards";
import * as React from "react";
import { useTranslation } from "react-i18next";
import { useParams } from "react-router-dom";
import { Button, Field, Image, Label, Spinner, Text } from "@fluentui/react-components";
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
      const renderCard = adaptiveCard.render();
      if (renderCard) {
        document.getElementsByClassName("card-area")[0].appendChild(renderCard);
      }
      adaptiveCard.onExecuteAction = function (action: any) {
        window.open(action.url, "_blank");
      };
      setLoader(false);
    }
  }, [isCardReady, consentState.isConsentsUpdated, messageState.isDraftMsgUpdated]);

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
        resultedTeams.push(
          <li>
            <Image src={ImageUtil.makeInitialImage(element)} />
            <span style={{ verticalAlign: "top", paddingLeft: "5px" }}>{element}</span>
          </li>
        );
      });
    }
    return resultedTeams;
  };

  const renderAudienceSelection = () => {
    if (consentState.teamNames && consentState.teamNames.length > 0) {
      return (
        <div key="teamNames">
          <Label>{t("TeamsLabel")}</Label>
          <ul style={{ listStyleType: "none" }}>{getItemList(consentState.teamNames)}</ul>
        </div>
      );
    } else if (consentState.rosterNames && consentState.rosterNames.length > 0) {
      return (
        <div key="rosterNames">
          <Label>{t("TeamsMembersLabel")}</Label>
          <ul style={{ listStyleType: "none" }}>{getItemList(consentState.rosterNames)}</ul>
        </div>
      );
    } else if (consentState.groupNames && consentState.groupNames.length > 0) {
      return (
        <div key="groupNames">
          <Label>{t("GroupsMembersLabel")}</Label>
          <ul style={{ listStyleType: "none" }}>{getItemList(consentState.groupNames)}</ul>
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
      {loader && <Spinner />}
      <>
        <div className="adaptive-task-grid">
          <div className="form-area">
            {!loader && (
              <>
                <Field size="large" label={t("ConfirmToSend")}>
                  <Text>{t("SendToRecipientsLabel")}</Text>
                </Field>
                <div style={{ margin: "16px" }}>{renderAudienceSelection()}</div>
              </>
            )}
          </div>
          <div className="card-area"></div>
        </div>
        <div className="fixed-footer">
          <div className="footer-actions">
            <div className="footer-button">
              <Button onClick={onSendMessage} appearance="primary">
                {t("Send")}
              </Button>
            </div>
          </div>
        </div>
      </>
    </>
  );
};
