// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TooltipHost } from 'office-ui-fabric-react';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import {
    Button, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Table, TableBody, TableCell,
    TableCellLayout, TableHeader, TableHeaderCell, TableRow, useArrowNavigationGroup
} from '@fluentui/react-components';
import {
    BookExclamationMark24Regular, CalendarCancel24Regular, CheckmarkSquare24Regular,
    DocumentCopyRegular, DocumentRegular, MoreHorizontal24Filled, OpenRegular,
    ShareScreenStop24Regular, Warning24Regular
} from '@fluentui/react-icons';
import * as microsoftTeams from '@microsoft/teams-js';
import { cancelSentNotification, duplicateDraftNotification } from '../../apis/messageListApi';
import { getBaseUrl } from '../../configVariables';
import { formatNumber } from '../../i18n';
import { ROUTE_PARTS, ROUTE_QUERY_PARAMS } from '../../routes';

export const SentMessageDetail = (sentMessages: any) => {
  const { t } = useTranslation();
  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });
  const statusUrl = (id: string) => getBaseUrl() + `/${ROUTE_PARTS.VIEW_STATUS}/${id}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}`;

  const columns = [
    { columnKey: "title", label: t("TitleText") },
    { columnKey: "status", label: "Status" },
    { columnKey: "recipients", label: t("Recipients") },
    { columnKey: "sent", label: t("Sent") },
    { columnKey: "createdBy", label: t("CreatedBy") },
  ];

  const renderSendingText = (message: any) => {
    var text = "";
    switch (message.status) {
      case "Queued":
        text = t("Queued");
        break;
      case "SyncingRecipients":
        text = t("SyncingRecipients");
        break;
      case "InstallingApp":
        text = t("InstallingApp");
        break;
      case "Sending":
        let sentCount =
          (message.succeeded ? message.succeeded : 0) +
          (message.failed ? message.failed : 0) +
          (message.unknown ? message.unknown : 0);
        text = t("SendingMessages", {
          SentCount: formatNumber(sentCount),
          TotalCount: formatNumber(message.totalMessageCount),
        });
        break;
      case "Canceling":
        text = t("Canceling");
        break;
      case "Canceled":
      case "Sent":
      case "Failed":
        text = "";
    }

    return text;
  };

  const shouldNotShowCancel = (msg: any) => {
    let cancelState = false;
    if (msg !== undefined && msg.status !== undefined) {
      const status = msg.status.toUpperCase();
      cancelState =
        status === "SENT" ||
        status === "UNKNOWN" ||
        status === "FAILED" ||
        status === "CANCELED" ||
        status === "CANCELING";
    }
    return cancelState;
  };

  const onOpenTaskModule = (event: any, url: string, title: string) => {
    let taskInfo: microsoftTeams.TaskInfo = {
      url: url,
      title: title,
      height: microsoftTeams.TaskModuleDimension.Medium,
      width: microsoftTeams.TaskModuleDimension.Medium,
      fallbackUrl: url,
    };
    let submitHandler = (err: any, result: any) => {};
    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
  };

  const duplicateDraftMessage = async (id: number) => {
    try {
      await duplicateDraftNotification(id);
    } catch (error) {
      return error;
    }
  };

  const cancelSentMessage = async (id: number) => {
    try {
      await cancelSentNotification(id);
    } catch (error) {
      return error;
    }
  };

  return (
    <Table {...keyboardNavAttr} role="grid" aria-label="Table with grid keyboard navigation">
      <TableHeader>
        <TableRow>
          {columns.map((column) => (
            <TableHeaderCell key={column.columnKey}>
              <b>{column.label}</b>
            </TableHeaderCell>
          ))}
          <TableHeaderCell key="actions" style={{ float: "right" }}>
            <b>Actions</b>
          </TableHeaderCell>
        </TableRow>
      </TableHeader>
      <TableBody>
        {sentMessages!.sentMessages!.map((item: any) => (
          <TableRow>
            <TableCell tabIndex={0} role="gridcell">
              <TableCellLayout media={<DocumentRegular />}>{item.title}</TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role="gridcell">
              <TableCellLayout>{renderSendingText(item)}</TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role="gridcell">
              <TableCellLayout>
                <div>
                  {/* <TooltipHost content={t("TooltipSuccess")} calloutProps={{ gapSpace: 0 }}> */}
                  <CheckmarkSquare24Regular style={{ color: "green", verticalAlign: "middle" }} />
                  <span
                    style={{ verticalAlign: "middle", paddingLeft: "2px", paddingRight: "8px" }}
                    className="semiBold"
                  >
                    {formatNumber(item.succeeded)}
                  </span>
                  {/* </TooltipHost>
                      <TooltipHost content={t("TooltipFailure")} calloutProps={{ gapSpace: 0 }}> */}
                  {/* <CloseIcon xSpacing="both" className="failed" outline /> */}
                  <ShareScreenStop24Regular style={{ color: "red", verticalAlign: "middle" }} />
                  <span
                    style={{ verticalAlign: "middle", paddingLeft: "2px", paddingRight: "8px" }}
                    className="semiBold"
                  >
                    {formatNumber(item.failed)}
                  </span>
                  {/* </TooltipHost> */}
                  {item.canceled && (
                    <TooltipHost content="Canceled" calloutProps={{ gapSpace: 0 }}>
                      <BookExclamationMark24Regular style={{ color: "yellow", verticalAlign: "middle" }} />
                      <span
                        style={{ verticalAlign: "middle", paddingLeft: "2px", paddingRight: "8px" }}
                        className="semiBold"
                      >
                        {formatNumber(item.canceled)}
                      </span>
                    </TooltipHost>
                  )}
                  {item.unknown && (
                    <TooltipHost content="Unknown" calloutProps={{ gapSpace: 0 }}>
                      <Warning24Regular style={{ color: "orange", verticalAlign: "middle" }} />
                      <span
                        style={{ verticalAlign: "middle", paddingLeft: "2px", paddingRight: "8px" }}
                        className="semiBold"
                      >
                        {formatNumber(item.unknown)}
                      </span>
                    </TooltipHost>
                  )}
                </div>
              </TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role="gridcell">
              <TableCellLayout>{item.sentDate}</TableCellLayout>
            </TableCell>
            <TableCell tabIndex={0} role="gridcell" style={{ textOverflow: "ellipsis" }}>
              <TableCellLayout>{item.createdBy}</TableCellLayout>
            </TableCell>
            <TableCell role="gridcell">
              <TableCellLayout style={{ float: "right" }}>
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <Button icon={<MoreHorizontal24Filled />} />
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem
                        icon={<OpenRegular />}
                        onClick={() => onOpenTaskModule(null, statusUrl(item.id), t("ViewStatus"))}
                      >
                        {t("ViewStatus")}
                      </MenuItem>
                      <MenuItem icon={<DocumentCopyRegular />} onClick={() => duplicateDraftMessage(item.id)}>
                        {t("Duplicate")}
                      </MenuItem>
                      <MenuItem
                        icon={<CalendarCancel24Regular />}
                        onClick={() => cancelSentMessage(item.id)}
                        disabled={shouldNotShowCancel(item)}
                      >
                        {t("Cancel")}
                      </MenuItem>
                    </MenuList>
                  </MenuPopover>
                </Menu>
              </TableCellLayout>
            </TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
};
