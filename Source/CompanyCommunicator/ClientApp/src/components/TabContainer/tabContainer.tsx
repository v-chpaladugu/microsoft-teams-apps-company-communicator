// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./tabContainer.scss";

import * as React from "react";
import { useTranslation, WithTranslation } from "react-i18next";

import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  Link,
  teamsDarkTheme,
} from "@fluentui/react-components";
import { ChatMultiple24Regular, PersonFeedback24Regular, QuestionCircle24Regular } from "@fluentui/react-icons";
import * as microsoftTeams from "@microsoft/teams-js";

import { getBaseUrl } from "../../configVariables";
import { GetFluentUITheme } from "../../constants";
import { DraftMessages } from "../DraftMessages/draftMessages";
import { Messages } from "../Messages/messages";
import { TableTest } from "../testData";

interface ITaskInfo {
  title?: string;
  height?: number;
  width?: number;
  url?: string;
  card?: string;
  fallbackUrl?: string;
  completionBotId?: string;
}

export interface ITaskInfoProps extends WithTranslation {
  getDraftMessagesList?: any;
}

export interface ITabContainerState {
  url: string;
}

export const TabContainer = (taskInfoPros: ITaskInfoProps, tabState: ITabContainerState) => {
  // readonly localize: TFunction;
  const url = getBaseUrl() + "/newmessage?locale={locale}";
  const { t } = useTranslation();
  const theme = GetFluentUITheme();

  React.useEffect(() => {
    //- Handle the Esc key
    document.addEventListener("keydown", escFunction, false);
  }, []);

  const escFunction = (event: any) => {
    if (event.keyCode === 27 || event.key === "Escape") {
      microsoftTeams.tasks.submitTask();
    }
  };

  const onNewMessage = () => {
    let taskInfo: ITaskInfo = {
      url,
      title: t("NewMessage"),
      height: 530,
      width: 1000,
      fallbackUrl: url,
    };

    let submitHandler = (err: any, result: any) => {
      taskInfoPros.getDraftMessagesList();
    };

    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
  };

  const buttonId = "callout-button";
  const customHeaderImagePath = process.env.REACT_APP_HEADERIMAGE;
  const customHeaderText =
    process.env.REACT_APP_HEADERTEXT == null ? t("CompanyCommunicator") : t(process.env.REACT_APP_HEADERTEXT);

  return (
    <>
      <div className={theme === teamsDarkTheme ? "company-communicator-header-dark" : "company-communicator-header"}>
        <div className="cc-main-heading">
          <img
            src={
              customHeaderImagePath == null ? require("../../assets/Images/mslogo.png").default : customHeaderImagePath
            }
            alt="Ms Logo"
            className="ms-logo"
            title={customHeaderText}
          />
          <span className="cc-header-text" title={customHeaderText}>
            {customHeaderText}
          </span>
        </div>
        <div className="cc-main-icons">
          <span className="cc-icon-span">
            <Link title={t("Support")} className="cc-link-icon" target="_blank" href="https://aka.ms/M365CCIssues">
              <QuestionCircle24Regular className="cc-icon" />
            </Link>
          </span>
          <span className="cc-icon-span">
            <Link title={t("Feedback")} className="cc-link-icon" target="_blank" href="https://aka.ms/M365CCFeedback">
              <PersonFeedback24Regular className="cc-icon" />
            </Link>
          </span>
        </div>
      </div>
      <div className="new-message">
        <Button icon={<ChatMultiple24Regular />} appearance="primary" onClick={onNewMessage}>
          {t("NewMessage")}
        </Button>
      </div>
      <div>
        <Accordion defaultOpenItems={["1", "2"]} multiple collapsible>
          <AccordionItem value="1">
            <AccordionHeader>{t("DraftMessagesSectionTitle")}</AccordionHeader>
            <AccordionPanel style={{ paddingTop: "16px", paddingBottom: "16px" }}>
              <div>
                <DraftMessages />
              </div>
            </AccordionPanel>
          </AccordionItem>
          <AccordionItem value="2">
            <AccordionHeader>{t("SentMessagesSectionTitle")}</AccordionHeader>
            <AccordionPanel style={{ paddingTop: "16px", paddingBottom: "16px" }}>
              <div>
                <Messages />
              </div>
            </AccordionPanel>
          </AccordionItem>
        </Accordion>
      </div>
      {/* <div className="cc-header">
              <Flex gap="gap.small" space="between">
                <Flex gap="gap.small" vAlign="center">
                  <img
                    src={
                      customHeaderImagePath == null
                        ? require("../../assets/Images/mslogo.png").default
                        : customHeaderImagePath
                    }
                    alt="Ms Logo"
                    className="ms-logo"
                    title={customHeaderText}
                  />
                  <span className="header-text" title={customHeaderText}>
                    {customHeaderText}
                  </span>
                </Flex>
                <Flex gap="gap.large" vAlign="center">
                  <FlexItem>
                    <a href="https://aka.ms/M365CCIssues" target="_blank" rel="noreferrer">
                      <Tooltip
                        trigger={
                          <img
                            src={require("../../assets/Images/HelpIcon.svg").default}
                            alt="Help"
                            className="support-icon"
                          />
                        }
                        content={this.localize("Support")}
                        pointing={false}
                      />
                    </a>
                  </FlexItem>
                  <FlexItem>
                    <a href="https://aka.ms/M365CCFeedback" target="_blank" rel="noreferrer">
                      <Tooltip
                        trigger={
                          <img
                            src={require("../../assets/Images/FeedbackIcon.svg").default}
                            alt="Feedback"
                            className="feedback-icon"
                          />
                        }
                        content={this.localize("Feedback")}
                        pointing={false}
                      />
                    </a>
                  </FlexItem>
                </Flex>
              </Flex>
            </div>
            <Flex className="tabContainer" column fill gap="gap.small">
              <Flex className="newPostBtn" hAlign="end" vAlign="end">
                <Button content={this.localize("NewMessage")} onClick={this.onNewMessage} primary />
              </Flex>
              <Flex className="messageContainer">
                <Flex.Item grow={1}>
                  <Accordion defaultActiveIndex={[0, 1]} panels={panels} />
                </Flex.Item>
              </Flex>
            </Flex> */}
    </>
  );
};
