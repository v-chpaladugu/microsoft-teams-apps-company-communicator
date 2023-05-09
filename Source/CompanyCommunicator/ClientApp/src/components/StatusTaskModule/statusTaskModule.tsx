// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import './statusTaskModule.scss';

import * as AdaptiveCards from 'adaptivecards';
import { TooltipHost } from 'office-ui-fabric-react';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { useParams } from 'react-router-dom';

import {
    AcceptIcon, Button, DownloadIcon, Flex, Image, List, Loader
} from '@fluentui/react-northstar';
import * as microsoftTeams from '@microsoft/teams-js';

import { exportNotification, getSentNotification } from '../../apis/messageListApi';
import { formatDate, formatDuration, formatNumber } from '../../i18n';
import { ImageUtil } from '../../utility/imageutility';
import {
    getInitAdaptiveCard, setCardAuthor, setCardBtn, setCardImageLink, setCardSummary, setCardTitle
} from '../AdaptiveCard/adaptiveCard';

export interface IListItem {
  header: string;
  media: JSX.Element;
}

export interface IMessage {
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
}

export interface IStatusState {
  message: IMessage;
  loader: boolean;
  page: string;
  teamId?: string;
}

export const StatusTaskModule = () => {
  const { t } = useTranslation();
  const {id} = useParams() as any;
  const [messageState, setMessageState] = React.useState<IStatusState>({
    message: { id: "", title: "" },
    loader: true,
    page: "ViewStatus",
    teamId: "",
  });

  let card: any;

  React.useEffect(() => {
    card = getInitAdaptiveCard(t);
    initialLoad();
  }, []);

  const initialLoad = () => {
    microsoftTeams.getContext((context) => {
      setMessageState({ ...messageState, teamId: context.teamId });
    });

    if (id) {
      getItem(id).then(() => {
        setMessageState({ ...messageState, loader: false });

        setCardTitle(card, messageState.message.title);
        setCardImageLink(card, messageState.message.imageLink);
        setCardSummary(card, messageState.message.summary);
        setCardAuthor(card, messageState.message.author);
        if (messageState.message.buttonTitle !== "" && messageState.message.buttonLink !== "") {
          setCardBtn(card, messageState.message.buttonTitle, messageState.message.buttonLink);
        }

        let adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        let renderedCard = adaptiveCard.render();
        document.getElementsByClassName("adaptiveCardContainer")[0].appendChild(renderedCard);
        let link = messageState.message.buttonLink;
        adaptiveCard.onExecuteAction = function (action) {
          window.open(link, "_blank");
        };
      });
    }
  };

  const getItem = async (id: number) => {
    try {
      const response = await getSentNotification(id);
      response.data.sendingDuration = formatDuration(response.data.sendingStartedDate, response.data.sentDate);
      response.data.sendingStartedDate = formatDate(response.data.sendingStartedDate);
      response.data.sentDate = formatDate(response.data.sentDate);
      response.data.succeeded = formatNumber(response.data.succeeded);
      response.data.failed = formatNumber(response.data.failed);
      response.data.unknown = response.data.unknown && formatNumber(response.data.unknown);
      response.data.canceled = response.data.canceled && formatNumber(response.data.canceled);
      // response.data.createdBy = response.data.createdBy;
      setMessageState({ ...messageState, message: response.data });
    } catch (error) {
      return error;
    }
  };

  const onClose = () => {
    microsoftTeams.tasks.submitTask();
  };

  const onExport = async () => {
    let spanner = document.getElementsByClassName("sendingLoader");
    spanner[0].classList.remove("hiddenLoader");
    let payload = {
      id: messageState.message.id,
      teamId: messageState.teamId,
    };
    await exportNotification(payload)
      .then(() => {
        setMessageState({ ...messageState, page: "SuccessPage" });
      })
      .catch(() => {
        setMessageState({ ...messageState, page: "ErrorPage" });
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
    if (messageState.message.teamNames && messageState.message.teamNames.length > 0) {
      return (
        <div>
          <h3>{t("SentToGeneralChannel")}</h3>
          <List items={getItemList(messageState.message.teamNames)} />
        </div>
      );
    } else if (messageState.message.rosterNames && messageState.message.rosterNames.length > 0) {
      return (
        <div>
          <h3>{t("SentToRosters")}</h3>
          <List items={getItemList(messageState.message.rosterNames)} />
        </div>
      );
    } else if (messageState.message.groupNames && messageState.message.groupNames.length > 0) {
      return (
        <div>
          <h3>{t("SentToGroups1")}</h3>
          <span>{t("SentToGroups2")}</span>
          <List items={getItemList(messageState.message.groupNames)} />
        </div>
      );
    } else if (messageState.message.allUsers) {
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
    if (messageState.message.errorMessage) {
      return (
        <div>
          <h3>{t("Errors")}</h3>
          <span>{messageState.message.errorMessage}</span>
        </div>
      );
    } else {
      return <div></div>;
    }
  };

  const renderWarningMessage = () => {
    if (messageState.message.warningMessage) {
      return (
        <div>
          <h3>{t("Warnings")}</h3>
          <span>{messageState.message.warningMessage}</span>
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
      {!messageState.loader && messageState.page === "ViewStatus" && (
        <div className="taskModule">
          <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
            <Flex className="scrollableContent">
              <Flex.Item size="size.half" className="formContentContainer">
                <Flex column>
                  <div className="contentField">
                    <h3>{t("TitleText")}</h3>
                    <span>{messageState.message.title}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("SendingStarted")}</h3>
                    <span>{messageState.message.sendingStartedDate}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Completed")}</h3>
                    <span>{messageState.message.sentDate}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Created By")}</h3>
                    <span>{messageState.message.createdBy}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Duration")}</h3>
                    <span>{messageState.message.sendingDuration}</span>
                  </div>
                  <div className="contentField">
                    <h3>{t("Results")}</h3>
                    <label>{t("Success", { SuccessCount: messageState.message.succeeded })}</label>
                    <br />
                    <label>{t("Failure", { FailureCount: messageState.message.failed })}</label>
                    {messageState.message.canceled && (
                      <>
                        <br />
                        <label>{t("Canceled", { CanceledCount: messageState.message.canceled })}</label>
                      </>
                    )}
                    {messageState.message.unknown && (
                      <>
                        <br />
                        <label>{t("Unknown", { UnknownCount: messageState.message.unknown })}</label>
                      </>
                    )}
                  </div>
                  <div className="contentField">{renderAudienceSelection()}</div>
                  <div className="contentField">{renderErrorMessage()}</div>
                  <div className="contentField">{renderWarningMessage()}</div>
                </Flex>
              </Flex.Item>
              <Flex.Item size="size.half">
                <div className="adaptiveCardContainer"></div>
              </Flex.Item>
            </Flex>
            <Flex className="footerContainer" vAlign="end" hAlign="end">
              <div className={messageState.message.canDownload ? "" : "disabled"}>
                <Flex className="buttonContainer" gap="gap.small">
                  <Flex.Item push>
                    <Loader
                      id="sendingLoader"
                      className="hiddenLoader sendingLoader"
                      size="smallest"
                      label={t("ExportLabel")}
                      labelPosition="end"
                    />
                  </Flex.Item>
                  <Flex.Item>
                    <TooltipHost
                      content={
                        !messageState.message.sendingCompleted
                          ? ""
                          : messageState.message.canDownload
                          ? ""
                          : t("ExportButtonProgressText")
                      }
                      calloutProps={{ gapSpace: 0 }}
                    >
                      <Button
                        icon={<DownloadIcon size="medium" />}
                        disabled={!messageState.message.canDownload || !messageState.message.sendingCompleted}
                        content={t("ExportButtonText")}
                        id="exportBtn"
                        onClick={onExport}
                        primary
                      />
                    </TooltipHost>
                  </Flex.Item>
                </Flex>
              </div>
            </Flex>
          </Flex>
        </div>
      )}
      {!messageState.loader && messageState.page === "SuccessPage" && (
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
                <Button content={t("CloseText")} id="closeBtn" onClick={onClose} primary />
              </Flex>
            </Flex>
          </Flex>
        </div>
      )}
      {!messageState.loader && messageState.page !== "ViewStatus" && messageState.page !== "SuccessPage" && (
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
                <Button content={t("CloseText")} id="closeBtn" onClick={onClose} primary />
              </Flex>
            </Flex>
          </Flex>
        </div>
      )}
    </>
  );
};
