// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import './mainContainer.scss';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import {
    Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Link
} from '@fluentui/react-components';
import {
    ChatMultiple24Regular, PersonFeedback24Regular, QuestionCircle24Regular
} from '@fluentui/react-icons';
import * as microsoftTeams from '@microsoft/teams-js';
import { GetDraftMessagesAction } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import { useAppDispatch } from '../../store';
import { DraftMessages } from '../DraftMessages/draftMessages';
import { Messages } from '../Messages/messages';

interface ITaskInfo {
  title?: string;
  height?: number;
  width?: number;
  url?: string;
  card?: string;
  fallbackUrl?: string;
  completionBotId?: string;
}

export const MainContainer = () => {
  const url = getBaseUrl() + "/newmessage?locale={locale}";
  const { t } = useTranslation();
  const dispatch = useAppDispatch();

  React.useEffect(() => {
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
      GetDraftMessagesAction(dispatch);
    };

    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
  };

  const customHeaderImagePath = process.env.REACT_APP_HEADERIMAGE;
  const customHeaderText =
    process.env.REACT_APP_HEADERTEXT == null ? t("CompanyCommunicator") : t(process.env.REACT_APP_HEADERTEXT);

  return (
    <>
      <div className="cc-header">
        <div className="cc-main-left">
          <img
            src={
              customHeaderImagePath == null ? require("../../assets/Images/mslogo.png").default : customHeaderImagePath
            }
            alt="Microsoft logo"
            className="cc-logo"
            title={customHeaderText}
          />
          <span className="cc-title" title={customHeaderText}>
            {customHeaderText}
          </span>
        </div>
        <div className="cc-main-right">
          <span className="cc-icon-holder">
            <Link title={t("Support")} className="cc-icon-link" target="_blank" href="https://aka.ms/M365CCIssues">
              <QuestionCircle24Regular className="cc-icon" />
            </Link>
          </span>
          <span className="cc-icon-holder">
            <Link title={t("Feedback")} className="cc-icon-link" target="_blank" href="https://aka.ms/M365CCFeedback">
              <PersonFeedback24Regular className="cc-icon" />
            </Link>
          </span>
        </div>
      </div>
      <div className="cc-new-message">
        <Button icon={<ChatMultiple24Regular />} appearance="primary" onClick={onNewMessage}>
          {t("NewMessage")}
        </Button>
      </div>
      <Accordion defaultOpenItems={["1", "2"]} multiple collapsible>
        <AccordionItem value="1">
          <AccordionHeader>{t("DraftMessagesSectionTitle")}</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            <DraftMessages />
          </AccordionPanel>
        </AccordionItem>
        <AccordionItem value="2">
          <AccordionHeader>{t("SentMessagesSectionTitle")}</AccordionHeader>
          <AccordionPanel className="cc-accordion-panel">
            <Messages />
          </AccordionPanel>
        </AccordionItem>
      </Accordion>
    </>
  );
};
