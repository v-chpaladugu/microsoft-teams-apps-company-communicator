import {
    TableBody,
    TableCell,
    TableRow,
    Table,
    TableHeader,
    TableHeaderCell,
    TableCellLayout,
    PresenceBadgeStatus,
    Avatar,
    Button,
    useArrowNavigationGroup,
    useFluent,
    useScrollbarWidth
  } from '@fluentui/react-components';
  import * as React from 'react';
  import {
    FolderRegular,
    EditRegular,
    OpenRegular,
    DocumentRegular,
    PeopleRegular,
    DocumentPdfRegular,
    VideoRegular,
    MoreHorizontal24Filled,
    DocumentHeader24Regular,
    MailMultiple24Regular,
    Checkmark24Filled,
    CheckmarkSquare24Regular,
    ShareScreenStop24Regular
  } from '@fluentui/react-icons';
  
  import { Menu, MenuTrigger, MenuList, MenuItem, MenuPopover } from '@fluentui/react-components';
import { useTranslation } from 'react-i18next';
import { formatNumber } from '../../i18n';
import { TooltipHost } from 'office-ui-fabric-react';
  

  export const SentMessageDetail = (sentMessages: any) => {
    const { t } = useTranslation();
    const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
    // let dd = new Array(draftMessages.length);
    // dd = [...draftMessages];
   // const { targetDocument } = useFluent();
   // const scrollbarWidth = useScrollbarWidth({ targetDocument: window.document });

    const columns = [
      { columnKey: 'title', label: t("TitleText") },
      { columnKey: 'status', label: "Status" },
      { columnKey: 'Recipients', label: t("Recipients") },
      { columnKey: 'Sent', label: t("Sent") },
      { columnKey: 'CreatedBy', label: t("CreatedBy") }
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

              text = t("SendingMessages", { "SentCount": formatNumber(sentCount), "TotalCount": formatNumber(message.totalMessageCount) });
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
  }

    return (
      <Table {...keyboardNavAttr} role='grid' aria-label='Table with grid keyboard navigation'>
        <TableHeader>
          <TableRow>
            {columns.map((column) => (
              <TableHeaderCell key={column.columnKey}><b>{column.label}</b></TableHeaderCell>
            ))}  
            <TableHeaderCell />
          </TableRow>
        </TableHeader>
        <TableBody>
          {sentMessages!.sentMessages!.map((item: any) => (
            <TableRow>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout media={<DocumentRegular />}>{item.title}</TableCellLayout>
              </TableCell>
               <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout>{renderSendingText(item)}</TableCellLayout>
              </TableCell>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout>
                <div>
                      {/* <TooltipHost content={t("TooltipSuccess")} calloutProps={{ gapSpace: 0 }}> */}
                          <CheckmarkSquare24Regular style={{color:"green", verticalAlign: 'middle'}}/>
                          <span style={{verticalAlign: 'middle', paddingLeft: '2px',  paddingRight: '8px'}} className="semiBold">{formatNumber(item.succeeded)}</span>
                      {/* </TooltipHost>
                      <TooltipHost content={t("TooltipFailure")} calloutProps={{ gapSpace: 0 }}> */}
                          {/* <CloseIcon xSpacing="both" className="failed" outline /> */}
                          <ShareScreenStop24Regular style={{color:"red", verticalAlign: 'middle'}}/>
                          <span style={{verticalAlign: 'middle', paddingLeft: '2px',  paddingRight: '8px'}} className="semiBold">{formatNumber(item.failed)}</span>
                      {/* </TooltipHost> */}
                      {
                          item.canceled &&
                          <TooltipHost content="Canceled" calloutProps={{ gapSpace: 0 }}>
                              {/* <ExclamationCircleIcon xSpacing="both" className="canceled" outline /> */}
                              <span className="semiBold">{formatNumber(item.canceled)}</span>
                          </TooltipHost>
                      }
                      {
                          item.unknown &&
                          <TooltipHost content="Unknown" calloutProps={{ gapSpace: 0 }}>
                              {/* <ExclamationTriangleIcon xSpacing="both" className="unknown" outline /> */}
                              <span className="semiBold">{formatNumber(item.unknown)}</span>
                          </TooltipHost>
                      }
                  </div>
                </TableCellLayout>
              </TableCell>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout>{item.sentDate}</TableCellLayout>
              </TableCell>
              <TableCell tabIndex={0} role='gridcell' style={{textOverflow: 'ellipsis'}}>
                <TableCellLayout>{item.createdBy}</TableCellLayout>
              </TableCell>
              <TableCell role='gridcell'>
                <TableCellLayout style={{float: 'right'}}>
                  <Menu>
                    <MenuTrigger disableButtonEnhancement>
                      <Button icon={<MoreHorizontal24Filled />} />
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        <MenuItem>New </MenuItem>
                        <MenuItem>New Window</MenuItem>
                        <MenuItem disabled>Open File</MenuItem>
                        <MenuItem>Open Folder</MenuItem>
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
  