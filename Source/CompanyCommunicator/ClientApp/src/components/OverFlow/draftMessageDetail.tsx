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
    MailMultiple24Regular
  } from '@fluentui/react-icons';
  
  import { Menu, MenuTrigger, MenuList, MenuItem, MenuPopover } from '@fluentui/react-components';
import { useTranslation } from 'react-i18next';
  

  export const DraftMessageDetail = (draftMessages: any) => {
    const { t } = useTranslation();
    const keyboardNavAttr = useArrowNavigationGroup({ axis: 'grid' });
    // let dd = new Array(draftMessages.length);
    // dd = [...draftMessages];

    const columns = [
      { columnKey: 'title', label: t("TitleText") }
    ];
  
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
          {draftMessages!.draftMessages!.map((item: any) => (
            <TableRow>
              <TableCell tabIndex={0} role='gridcell'>
                <TableCellLayout media={<DocumentRegular />}>{item.title}</TableCellLayout>
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
  