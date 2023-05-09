// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useTranslation } from 'react-i18next';

import { Spinner } from '@fluentui/react-components';

import { GetDraftMessagesAction } from '../../actions';
import { RootState, useAppDispatch, useAppSelector } from '../../store';
import { DraftMessageDetail } from '../OverFlow/draftMessageDetail';

export const DraftMessages = () => {
  const { t } = useTranslation();
  const draftMessages = useAppSelector((state: RootState) => state.messages).draftMessages.payload;
  const [loader, setLoader] = React.useState(true);
  const dispatch = useAppDispatch();

  React.useEffect(() => {
    GetDraftMessagesAction(dispatch);
  }, []);

  React.useEffect(() => {
    if (draftMessages && draftMessages.length > 0) {
      setLoader(false);
    }
  }, [draftMessages]);

  //   const label = t("TitleText");
  //   const outList = draftMessages.map(processItem);
  //   const allDraftMessages = [...label, ...outList];

  //   const processItem = (message: any) => {
  //     keyCount++;
  //     const out = {
  //       key: keyCount,
  //       content: (
  //         <Flex vAlign="center" fill gap="gap.small">
  //           <Flex.Item shrink={0} grow={1}>
  //             <Text>{message.title}</Text>
  //           </Flex.Item>
  //           <Flex.Item shrink={0} align="end">
  //             <Overflow message={message} title="" />
  //           </Flex.Item>
  //         </Flex>
  //       ),
  //       styles: { margin: "0.2rem 0.2rem 0 0" },
  //       onClick: (): void => {
  //         let url = getBaseUrl() + "/newmessage/" + message.id + "?locale={locale}";
  //         this.onOpenTaskModule(null, url, this.localize("EditMessage"));
  //       },
  //     };
  //     return out;
  //   };

  return (
    <>
      {loader && <Spinner labelPosition="below" size="large" label="Fetching draft messages..." />}
      {draftMessages && draftMessages.length === 0 && !loader && (
        <div className="results">{t("EmptyDraftMessages")}</div>
      )}
      {draftMessages && draftMessages.length > 0 && !loader && <DraftMessageDetail draftMessages={draftMessages} />}
    </>
  );
};
