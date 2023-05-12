// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./statusTaskModule.scss";

import * as AdaptiveCards from "adaptivecards";
import * as React from "react";
import { useTranslation } from "react-i18next";
import { useParams } from "react-router-dom";

import * as microsoftTeams from "@microsoft/teams-js";

import { exportNotification, getSentNotification } from "../../apis/messageListApi";
import { formatDate, formatDuration, formatNumber } from "../../i18n";
import { ImageUtil } from "../../utility/imageutility";

import { Button, MenuItem, MenuList, Image, Spinner } from "@fluentui/react-components";

import { ArrowDownload24Regular, CheckmarkSquare24Regular } from "@fluentui/react-icons";

import {
  getInitAdaptiveCard,
  setCardAuthor,
  setCardBtn,
  setCardImageLink,
  setCardSummary,
  setCardTitle,
} from "../AdaptiveCard/adaptiveCard";

export interface IListItem {
  header: string;
  media: JSX.Element;
}

export interface IMessageState {
  id: string;
  title: string;
  acknowledgements?: string;
  reactions?: string;
  responses?: string;
  succeeded?: string;
  failed?: string;
  unknown?: string;
  canceled?: string;
  sentDate?: string;
  imageLink?: string;
  summary?: string;
  author?: string;
  buttonLink?: string;
  buttonTitle?: string;
  teamNames?: string[];
  rosterNames?: string[];
  groupNames?: string[];
  allUsers?: boolean;
  sendingStartedDate?: string;
  sendingDuration?: string;
  errorMessage?: string;
  warningMessage?: string;
  canDownload?: boolean;
  sendingCompleted?: boolean;
  createdBy?: string;

  isMsgDataUpdated: boolean;
}

export interface IStatusState {
  page: string;
  teamId?: string;
  isTeamDataUpdated: boolean;
}

let card: any;

export const StatusTaskModule = () => {
  const { t } = useTranslation();
  const { id } = useParams() as any;
  const [loader, setLoader] = React.useState(true);
  const [isCardReady, setIsCardReady] = React.useState(false);

  const [messageState, setMessageState] = React.useState<IMessageState>({
    id: "",
    title: "",
    isMsgDataUpdated: false,
  });

  const [statusState, setStatusState] = React.useState<IStatusState>({
    page: "ViewStatus",
    teamId: "",
    isTeamDataUpdated: false,
  });

  React.useEffect(() => {
    microsoftTeams.getContext((context) => {
      setStatusState({ ...statusState, teamId: context.teamId, isTeamDataUpdated: true });
    });
  }, []);

  React.useEffect(() => {
    if (id) {
      getMessage(id);
    }
  }, [id]);

  React.useEffect(() => {
    if (isCardReady && messageState.isMsgDataUpdated) {
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
  }, [isCardReady, messageState.isMsgDataUpdated]);

  const getMessage = async (id: number) => {
    try {
      await getSentNotification(id).then((response) => {
        updateCardData(response.data);
        response.data.sendingDuration = formatDuration(response.data.sendingStartedDate, response.data.sentDate);
        response.data.sendingStartedDate = formatDate(response.data.sendingStartedDate);
        response.data.sentDate = formatDate(response.data.sentDate);
        response.data.succeeded = formatNumber(response.data.succeeded);
        response.data.failed = formatNumber(response.data.failed);
        response.data.unknown = response.data.unknown && formatNumber(response.data.unknown);
        response.data.canceled = response.data.canceled && formatNumber(response.data.canceled);
        setMessageState({ ...response.data, isMsgDataUpdated: true });
      });
    } catch (error) {
      return error;
    }
  };

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

  const onClose = () => {
    microsoftTeams.tasks.submitTask();
  };

  const onExport = async () => {
    let spanner = document.getElementsByClassName("sendingLoader");
    spanner[0].classList.remove("hiddenLoader");
    let payload = {
      id: messageState.id,
      teamId: statusState.teamId,
    };
    await exportNotification(payload)
      .then(() => {
        setStatusState({ ...statusState, page: "SuccessPage" });
      })
      .catch(() => {
        setStatusState({ ...statusState, page: "ErrorPage" });
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
    if (messageState.teamNames && messageState.teamNames.length > 0) {
      return (
        <div>
          <h3>{t("SentToGeneralChannel")}</h3>
          <MenuList>{getItemList(messageState.teamNames)}</MenuList>
        </div>
      );
    } else if (messageState.rosterNames && messageState.rosterNames.length > 0) {
      return (
        <div>
          <h3>{t("SentToRosters")}</h3>
          <MenuList>{getItemList(messageState.rosterNames)}</MenuList>
        </div>
      );
    } else if (messageState.groupNames && messageState.groupNames.length > 0) {
      return (
        <div>
          <h3>{t("SentToGroups1")}</h3>
          <span>{t("SentToGroups2")}</span>
          <MenuList>{getItemList(messageState.groupNames)}</MenuList>
        </div>
      );
    } else if (messageState.allUsers) {
      return (
        <div>
          <h3>{t("SendToAllUsers")}</h3>
        </div>
      );
    } else {
      return <div></div>;
    }
  };

  const renderErrorMessage = () => {
    if (messageState.errorMessage) {
      return (
        <div>
          <h3>{t("Errors")}</h3>
          <span>{messageState.errorMessage}</span>
        </div>
      );
    } else {
      return <div></div>;
    }
  };

  const renderWarningMessage = () => {
    if (messageState.warningMessage) {
      return (
        <div>
          <h3>{t("Warnings")}</h3>
          <span>{messageState.warningMessage}</span>
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
      {statusState.page === "ViewStatus" && (
        <>
          <div className="adaptive-task-grid">
            <div className="form-area">
              {!loader && (
                <>
                  <div className="contentField">
                    <h3>{t("TitleText")}</h3>
                    <span>{messageState.title}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("SendingStarted")}</h3>
                    <span>{messageState.sendingStartedDate}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Completed")}</h3>
                    <span>{messageState.sentDate}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Created By")}</h3>
                    <span>{messageState.createdBy}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Duration")}</h3>
                    <span>{messageState.sendingDuration}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Results")}</h3>
                    <label>{t("Success", { SuccessCount: messageState.succeeded })}</label>
                    <br />
                    <label>{t("Failure", { FailureCount: messageState.failed })}</label>
                    {messageState.canceled && (
                      <>
                        <br />
                        <label>{t("Canceled", { CanceledCount: messageState.canceled })}</label>
                      </>
                    )}
                    {messageState.unknown && (
                      <>
                        <br />
                        <label>{t("Unknown", { UnknownCount: messageState.unknown })}</label>
                      </>
                    )}
                  </div>
                  <div className="contentField">{renderAudienceSelection()}</div>
                  <div className="contentField">{renderErrorMessage()}</div>
                  <div className="contentField">{renderWarningMessage()}</div>
                </>
              )}
            </div>
            <div className="card-area"></div>
          </div>
          <div className="fixed-footer">
            <Spinner
              id="sendingLoader"
              className="spinner-wheel"
              size="small"
              label={t("ExportLabel")}
              labelPosition="before"
            />
            <Button
              icon={<ArrowDownload24Regular />}
              disabled={!messageState.canDownload || !messageState.sendingCompleted}
              id="exportBtn"
              onClick={onExport}
              appearance="primary"
            >
              {t("ExportButtonText")}
            </Button>
          </div>
        </>
      )}
      {!loader && statusState.page === "SuccessPage" && (
        <div className="taskModule">
          <div className="displayMessageField">
            <br />
            <br />
            <div>
              <span>
                <CheckmarkSquare24Regular />
              </span>
              <h1>{t("ExportQueueTitle")}</h1>
            </div>
            <span>{t("ExportQueueSuccessMessage1")}</span>
            <br />
            <br />
            <span>{t("ExportQueueSuccessMessage2")}</span>
            <br />
            <span>{t("ExportQueueSuccessMessage3")}</span>
          </div>
          <Button id="closeBtn" onClick={onClose} appearance="primary">
            {t("CloseText")}
          </Button>
        </div>
      )}
      {!loader && statusState.page !== "ViewStatus" && statusState.page !== "SuccessPage" && (
        <div className="taskModule">
          <div className="displayMessageField">
            <br />
            <br />
            <div>
              <span></span>
              <h1 className="light">{t("ExportErrorTitle")}</h1>
            </div>
            <span>{t("ExportErrorMessage")}</span>
          </div>
          <Button id="closeBtn" onClick={onClose} appearance="primary">
            {t("CloseText")}
          </Button>
        </div>
      )}
    </>
  );
};
