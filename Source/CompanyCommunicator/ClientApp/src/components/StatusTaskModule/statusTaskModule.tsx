// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./statusTaskModule.scss";

import * as AdaptiveCards from "adaptivecards";
import { TooltipHost } from "office-ui-fabric-react";
import * as React from "react";
import { useTranslation } from "react-i18next";
import { useParams } from "react-router-dom";

import { AcceptIcon, DownloadIcon, Flex } from "@fluentui/react-northstar";

import * as microsoftTeams from "@microsoft/teams-js";

import { exportNotification, getSentNotification } from "../../apis/messageListApi";
import { formatDate, formatDuration, formatNumber } from "../../i18n";
import { ImageUtil } from "../../utility/imageutility";

import { Button, MenuItem, MenuList, Text, Image, Spinner } from "@fluentui/react-components";

import parse from 'html-react-parser';

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
let renderCard: any;

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
      renderCard = adaptiveCard.render();
      if (messageState.buttonLink) {
        let link = messageState.buttonLink;
        adaptiveCard.onExecuteAction = function (action) {
          window.open(link, "_blank");
        };
      }
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
      {!loader && statusState.page === "ViewStatus" && (
        <div className="taskModule">
          <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
            <Flex className="scrollableContent">
              <Flex.Item size="size.half" className="formContentContainer">
                <Flex column>
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
                </Flex>
              </Flex.Item>
              <Flex.Item size="size.half">
                <div className="adaptiveCardContainer">{parse(renderCard.outerHTML)}</div>
              </Flex.Item>
            </Flex>
            <Flex className="footerContainer" vAlign="end" hAlign="end">
              <div className={messageState.canDownload ? "" : "disabled"}>
                <Flex className="buttonContainer" gap="gap.small">
                  <Flex.Item push>
                    <Spinner
                      id="sendingLoader"
                      className="hiddenLoader sendingLoader"
                      size="small"
                      label={t("ExportLabel")}
                      labelPosition="after"
                    />
                  </Flex.Item>
                  <Flex.Item>
                    <TooltipHost
                      content={
                        !messageState.sendingCompleted
                          ? ""
                          : messageState.canDownload
                          ? ""
                          : t("ExportButtonProgressText")
                      }
                      calloutProps={{ gapSpace: 0 }}
                    >
                      <Button
                        icon={<DownloadIcon size="medium" />}
                        disabled={!messageState.canDownload || !messageState.sendingCompleted}
                        id="exportBtn"
                        onClick={onExport}
                        appearance="primary"
                      >
                        {t("ExportButtonText")}
                      </Button>
                    </TooltipHost>
                  </Flex.Item>
                </Flex>
              </div>
            </Flex>
          </Flex>
        </div>
      )}
      {!loader && statusState.page === "SuccessPage" && (
        <div className="taskModule">
          <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
            <div className="displayMessageField">
              <br />
              <br />
              <div>
                <span>
                  <AcceptIcon className="iconStyle" xSpacing="before" size="largest" outline />
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
            <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
              <Flex className="buttonContainer">
                <Button id="closeBtn" onClick={onClose} appearance="primary">{t("CloseText")}</Button>
              </Flex>
            </Flex>
          </Flex>
        </div>
      )}
      {!loader && statusState.page !== "ViewStatus" && statusState.page !== "SuccessPage" && (
        <div className="taskModule">
          <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
            <div className="displayMessageField">
              <br />
              <br />
              <div>
                <span></span>
                <h1 className="light">{t("ExportErrorTitle")}</h1>
              </div>
              <span>{t("ExportErrorMessage")}</span>
            </div>
            <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
              <Flex className="buttonContainer">
                <Button id="closeBtn" onClick={onClose} appearance="primary">
                  {t("CloseText")}
                </Button>
              </Flex>
            </Flex>
          </Flex>
        </div>
      )}
    </>
  );
};
