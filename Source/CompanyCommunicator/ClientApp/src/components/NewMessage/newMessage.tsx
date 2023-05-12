// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./newMessage.scss";

import * as AdaptiveCards from "adaptivecards";
import * as React from "react";
import { useTranslation } from "react-i18next";
import { useParams } from "react-router-dom";

import {
  Button,
  Combobox,
  ComboboxProps,
  Field,
  Input,
  Label,
  makeStyles,
  Option,
  Radio,
  RadioGroup,
  shorthands,
  Spinner,
  Textarea,
  tokens,
  useId,
} from "@fluentui/react-components";
import { ArrowUpload24Regular, Dismiss12Regular } from "@fluentui/react-icons";
import * as microsoftTeams from "@microsoft/teams-js";

import { GetGroupsAction, GetTeamsDataAction, SearchGroupsAction, VerifyGroupAccessAction } from "../../actions";
import {
  createDraftNotification,
  getDraftNotification,
  searchGroups,
  updateDraftNotification,
} from "../../apis/messageListApi";
import { getBaseUrl } from "../../configVariables";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { ImageUtil } from "../../utility/imageutility";
import {
  getInitAdaptiveCard,
  setCardAuthor,
  setCardBtn,
  setCardImageLink,
  setCardSummary,
  setCardTitle,
} from "../AdaptiveCard/adaptiveCard";

const validImageTypes = ["image/gif", "image/jpeg", "image/png", "image/jpg"];

export interface formState {
  id?: string;
  title: string;
  imageLink?: string;
  summary?: string;
  author?: string;
  buttonTitle?: string;
  buttonLink?: string;
  teams: any[];
  rosters: any[];
  groups: any[];
  allUsers: boolean;
}

let card: any;

const useStyles = makeStyles({
  root: {
    // Stack the label above the field with a gap
    display: "grid",
    gridTemplateRows: "repeat(1fr)",
    justifyItems: "start",
    ...shorthands.gap("2px"),
    paddingLeft: "36px",
  },
  tagsList: {
    listStyleType: "none",
    marginBottom: tokens.spacingVerticalXXS,
    marginTop: 0,
    paddingLeft: 0,
    display: "flex",
    gridGap: tokens.spacingHorizontalXXS,
  },
});

enum AudienceSelection {
  Teams = "Teams",
  Rosters = "Rosters",
  Groups = "Groups",
  AllUsers = "AllUsers",
  None = "None",
}

enum CurrentPageSelection {
  CardCreation = "CardCreation",
  AudienceSelection = "AudienceSelection",
}

export const NewMessage = () => {
  let fileInput = React.createRef<any>();
  const MAX_SELECTED_TEAMS_NUM: number = 20;
  const { t } = useTranslation();
  const { id } = useParams() as any;
  const teams = useAppSelector((state: RootState) => state.messages).teamsData.payload;
  const groups = useAppSelector((state: RootState) => state.messages).groups.payload;
  const queryGroups = useAppSelector((state: RootState) => state.messages).queryGroups.payload;

  // const verifyGroupAccess = useAppSelector((state: RootState) => state.messages).verifyGroup.payload;

  const [selectedRadioButton, setSelectedRadioButton] = React.useState(AudienceSelection.None);
  const [pageSelection, setPageSelection] = React.useState(CurrentPageSelection.CardCreation);

  const [loader, setLoader] = React.useState(false);
  const [formState, setFormState] = React.useState<formState>({
    title: "",
    teams: [],
    rosters: [],
    groups: [],
    allUsers: false,
  });

  const dispatch = useAppDispatch();

  React.useEffect(() => {
    GetTeamsDataAction(dispatch);
    VerifyGroupAccessAction(dispatch);
    card = getInitAdaptiveCard(t);
    setDefaultCard(card);
    updateAdaptiveCard();
  }, []);

  React.useEffect(() => {
    if (id) {
      GetGroupsAction(dispatch, { id });
      getDraftNotificationItem(id);
    }
  }, [id]);

  const getDraftNotificationItem = async (id: number) => {
    try {
      await getDraftNotification(id).then((response) => {
        const draftMessageDetail = response.data;

        if (draftMessageDetail.teams.length > 0) {
          setSelectedRadioButton(AudienceSelection.Teams);
        } else if (draftMessageDetail.rosters.length > 0) {
          setSelectedRadioButton(AudienceSelection.Rosters);
        } else if (draftMessageDetail.groups.length > 0) {
          setSelectedRadioButton(AudienceSelection.Groups);
        } else if (draftMessageDetail.allUsers) {
          setSelectedRadioButton(AudienceSelection.AllUsers);
        }

        setFormState({
          ...formState,
          id: draftMessageDetail.id,
          title: draftMessageDetail.title,
          imageLink: draftMessageDetail.imageLink,
          summary: draftMessageDetail.summary,
          author: draftMessageDetail.author,
          buttonTitle: draftMessageDetail.buttonTitle,
          buttonLink: draftMessageDetail.buttonLink,
          teams: draftMessageDetail.teams,
          rosters: draftMessageDetail.rosters,
          groups: draftMessageDetail.groups,
          allUsers: draftMessageDetail.allUsers,
        });

        setCardTitle(card, draftMessageDetail.title);
        setCardImageLink(card, draftMessageDetail.imageLink);
        setCardSummary(card, draftMessageDetail.summary);
        setCardAuthor(card, draftMessageDetail.author);
        setCardBtn(card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink);

        updateAdaptiveCard();
      });
    } catch (error) {
      return error;
    }
  };

  const setDefaultCard = (card: any) => {
    const titleAsString = t("TitleText");
    const summaryAsString = t("Summary");
    const authorAsString = t("Author1");
    const buttonTitleAsString = t("ButtonTitle");
    setCardTitle(card, titleAsString);
    let imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
    setCardImageLink(card, imgUrl);
    setCardSummary(card, summaryAsString);
    setCardAuthor(card, authorAsString);
    setCardBtn(card, buttonTitleAsString, "https://adaptivecards.io");
  };

  const updateAdaptiveCard = () => {
    var adaptiveCard = new AdaptiveCards.AdaptiveCard();
    adaptiveCard.parse(card);
    const renderCard = adaptiveCard.render();
    if (renderCard) {
      document.getElementsByClassName("card-area")[0].innerHTML = "";
      document.getElementsByClassName("card-area")[0].appendChild(renderCard);
    }
    adaptiveCard.onExecuteAction = function (action: any) {
      window.open(action.url, "_blank");
    };
    setLoader(false);
  };

  const handleUploadClick = (event: any) => {
    if (fileInput.current) {
      fileInput.current.click();
    }
  };

  const checkValidSizeOfImage = (resizedImageAsBase64: string) => {
    var stringLength = resizedImageAsBase64.length - "data:image/png;base64,".length;
    var sizeInBytes = 4 * Math.ceil(stringLength / 3) * 0.5624896334383812;
    var sizeInKb = sizeInBytes / 1000;

    if (sizeInKb <= 1024) return true;
    else return false;
  };

  const handleImageSelection = () => {
    const file = fileInput.current.files[0];

    if (file) {
      const fileType = file["type"];
      const { type: mimeType } = file;

      if (!validImageTypes.includes(fileType)) {
        // setFormState({ ...formState, errorImageUrlMessage: t("ErrorImageTypesMessage") });
        return;
      }

      // setFormState({ ...formState, localImagePath: file["name"] });
      // setFormState({ ...formState, errorImageUrlMessage: "" });

      const fileReader = new FileReader();
      fileReader.readAsDataURL(file);
      fileReader.onload = () => {
        var image = new Image();
        image.src = fileReader.result as string;
        var resizedImageAsBase64 = fileReader.result as string;

        image.onload = function (e: any) {
          const MAX_WIDTH = 1024;

          if (image.width > MAX_WIDTH) {
            const canvas = document.createElement("canvas");
            canvas.width = MAX_WIDTH;
            canvas.height = ~~(image.height * (MAX_WIDTH / image.width));
            const context = canvas.getContext("2d", { alpha: false });
            if (!context) {
              return;
            }
            context.drawImage(image, 0, 0, canvas.width, canvas.height);
            resizedImageAsBase64 = canvas.toDataURL(mimeType);
          }
        };

        if (!checkValidSizeOfImage(resizedImageAsBase64)) {
          // setFormState({ ...formState, errorImageUrlMessage: t("ErrorImageSizeMessage") });
          return;
        }

        setCardImageLink(card, resizedImageAsBase64);
        // updateCard();
        setFormState({ ...formState, imageLink: resizedImageAsBase64 });
      };

      fileReader.onerror = (error) => {
        //reject(error);
      };
    }
  };

  const isSaveBtnDisabled = () => {
    return true;
  };

  const isNextBtnDisabled = () => {
    return false;
  };

  const onTeamsChange = (event: any, itemsData: any) => {
    setFormState({
      ...formState,
      teams: [],
      allUsers: false,
    });
  };

  const onRostersChange = (event: any, itemsData: any) => {
    setFormState({
      ...formState,
      rosters: [],
      allUsers: false,
    });
  };

  const onGroupsChange = (event: any, data: any) => {
    setFormState({
      ...formState,
      groups: [],
      allUsers: false,
    });
  };

  // encodeURIComponent(itemsData.searchQuery);

  const onSave = () => {
    // if (formState.exists) {
    //   editDraftMessage(draftMessage).then(() => {
    //     microsoftTeams.tasks.submitTask();
    //   });
    // } else {
    //   postDraftMessage(draftMessage).then(() => {
    //     microsoftTeams.tasks.submitTask();
    //   });
    // }
  };

  const editDraftMessage = async (draftMessage: formState) => {
    try {
      await updateDraftNotification(draftMessage);
    } catch (error) {
      return error;
    }
  };

  const postDraftMessage = async (draftMessage: formState) => {
    try {
      await createDraftNotification(draftMessage);
    } catch (error) {
      throw error;
    }
  };

  const onNext = (event: any) => {
    setPageSelection(CurrentPageSelection.AudienceSelection);
  };

  const onBack = (event: any) => {
    setPageSelection(CurrentPageSelection.CardCreation);
  };

  const onTitleChanged = (event: any) => {
    setCardTitle(card, event.target.value);
    setFormState({ ...formState, title: event.target.value });
    updateAdaptiveCard();
  };

  const onImageLinkChanged = (event: any) => {
    // let url = event.target.value.toLowerCase();
    // if (
    //   !(
    //     url === "" ||
    //     url.startsWith("https://") ||
    //     url.startsWith("data:image/png;base64,") ||
    //     url.startsWith("data:image/jpeg;base64,") ||
    //     url.startsWith("data:image/gif;base64,")
    //   )
    // ) {
    //   setFormState({ ...formState, errorImageUrlMessage: t("ErrorURLMessage") });
    // } else {
    //   setFormState({ ...formState, errorImageUrlMessage: "" });
    // }
    // let showDefaultCard =
    //   !formState.title &&
    //   !event.target.value &&
    //   !formState.summary &&
    //   !formState.author &&
    //   !formState.btnTitle &&
    //   !formState.btnLink;
    // setCardTitle(card, formState.title);
    // setCardImageLink(card, event.target.value);
    // setCardSummary(card, formState.summary);
    // setCardAuthor(card, formState.author);
    // setCardBtn(card, formState.btnTitle, formState.btnLink);
    // setFormState({ ...formState, imageLink: event.target.value, card: card });
    // if (showDefaultCard) {
    //   setDefaultCard(card);
    // }
    // updateCard();
  };

  const onSummaryChanged = (event: any) => {
    setCardSummary(card, event.target.value);
    setFormState({ ...formState, summary: event.target.value });
    updateAdaptiveCard();
  };

  const onAuthorChanged = (event: any) => {
    setCardAuthor(card, event.target.value);
    setFormState({ ...formState, author: event.target.value });
    updateAdaptiveCard();
  };

  const onBtnTitleChanged = (event: any) => {
    setCardBtn(card, event.target.value, formState.buttonLink);
    setFormState({ ...formState, buttonTitle: event.target.value });
    updateAdaptiveCard();
  };

  const onBtnLinkChanged = (event: any) => {
    setCardBtn(card, formState.buttonTitle, event.target.value);
    setFormState({ ...formState, buttonLink: event.target.value });
    updateAdaptiveCard();
  };

  // generate ids for handling labelling
  const teamsComboId = useId("teams-combo-multi");
  const teamsSelectedListId = `${teamsComboId}-selection`;

  const rostersComboId = useId("rosters-combo-multi");
  const rostersSelectedListId = `${rostersComboId}-selection`;

  const searchComboId = useId("search-combo-multi");
  const searchSelectedListId = `${searchComboId}-selection`;

  // refs for managing focus when removing tags
  const teamsSelectedListRef = React.useRef<HTMLUListElement>(null);
  const teamsComboboxInputRef = React.useRef<HTMLInputElement>(null);
  const rostersSelectedListRef = React.useRef<HTMLUListElement>(null);
  const rostersComboboxInputRef = React.useRef<HTMLInputElement>(null);
  const searchSelectedListRef = React.useRef<HTMLUListElement>(null);
  const searchComboboxInputRef = React.useRef<HTMLInputElement>(null);

  // Handle selectedOptions both when an option is selected or deselected in the Combobox,
  // and when an option is removed by clicking on a tag
  const [teamsSelectedOptions, setTeamsSelectedOptions] = React.useState<string[]>([]);
  const [rostersSelectedOptions, setRostersSelectedOptions] = React.useState<string[]>([]);

  const [searchSelectedOptions, setSearchSelectedOptions] = React.useState<string[]>([]);

  const onTeamsSelect: ComboboxProps["onOptionSelect"] = (event, data) => {
    setTeamsSelectedOptions(data.selectedOptions);
  };

  const onRostersSelect: ComboboxProps["onOptionSelect"] = (event, data) => {
    setRostersSelectedOptions(data.selectedOptions);
  };

  const onSearchSelect: ComboboxProps["onOptionSelect"] = (event, data) => {
    if (data.optionText && !searchSelectedOptions.find((x) => x === data.optionText)) {
      setSearchSelectedOptions([...searchSelectedOptions, data.optionText]);
    }
  };

  const onSearchChange = (event: any) => {
    if (event && event.target && event.target.value) {
      SearchGroupsAction(dispatch, { query: event.target.value });
    }
    // console.log(event.target.value);
  };

  const onTeamsTagClick = (option: string, index: number) => {
    // remove selected option
    setTeamsSelectedOptions(teamsSelectedOptions.filter((o) => o !== option));

    // focus previous or next option, defaulting to focusing back to the combo input
    const indexToFocus = index === 0 ? 1 : index - 1;
    const optionToFocus = teamsSelectedListRef.current?.querySelector(`#${teamsComboId}-remove-${indexToFocus}`);
    if (optionToFocus) {
      (optionToFocus as HTMLButtonElement).focus();
    } else {
      teamsComboboxInputRef.current?.focus();
    }
  };

  const onRostersTagClick = (option: string, index: number) => {
    // remove selected option
    setRostersSelectedOptions(rostersSelectedOptions.filter((o) => o !== option));

    // focus previous or next option, defaulting to focusing back to the combo input
    const indexToFocus = index === 0 ? 1 : index - 1;
    const optionToFocus = rostersSelectedListRef.current?.querySelector(`#${rostersComboId}-remove-${indexToFocus}`);
    if (optionToFocus) {
      (optionToFocus as HTMLButtonElement).focus();
    } else {
      rostersComboboxInputRef.current?.focus();
    }
  };

  const onSearchTagClick = (option: string, index: number) => {
    // remove selected option
    setSearchSelectedOptions(searchSelectedOptions.filter((o) => o !== option));

    // focus previous or next option, defaulting to focusing back to the combo input
    const indexToFocus = index === 0 ? 1 : index - 1;
    const optionToFocus = searchSelectedListRef.current?.querySelector(`#${searchComboId}-remove-${indexToFocus}`);
    if (optionToFocus) {
      (optionToFocus as HTMLButtonElement).focus();
    } else {
      searchComboboxInputRef.current?.focus();
    }
  };

  const teamsLabelledBy = teamsSelectedOptions.length > 0 ? `${teamsComboId} ${teamsSelectedListId}` : teamsComboId;
  const rostersLabelledBy =
    rostersSelectedOptions.length > 0 ? `${rostersComboId} ${rostersSelectedListId}` : rostersComboId;

  const searchLabelledBy =
    searchSelectedOptions.length > 0 ? `${searchComboId} ${searchSelectedListId}` : searchComboId;

  const styles = useStyles();

  return (
    <>
      {loader && <Spinner labelPosition="below" />}
      {!loader && (
        <div>
          {pageSelection === CurrentPageSelection.CardCreation && (
            <>
              <div className="adaptive-task-grid">
                <div className="form-area">
                  <Field size="large" label={t("TitleText")}>
                    <Input
                      placeholder={t("PlaceHolderTitle")}
                      onChange={onTitleChanged}
                      autoComplete="off"
                      size="large"
                      appearance="filled-darker"
                      value={formState.title}
                    />
                  </Field>
                  <Field size="large" label={t("ImageURL")}>
                    <div
                      style={{
                        display: "grid",
                        gridTemplateColumns: "1fr auto",
                        gridTemplateAreas: "inp-area btn-area",
                      }}
                    >
                      {/* <input
                      className="file-button"
                      aria-labelledby="imageLabelId"
                      type="file"
                      id="file"
                      name="file"
                      aria-label={t("ImageURL")}
                    /> */}
                      <Input
                        size="large"
                        style={{ gridColumn: "1" }}
                        appearance="filled-darker"
                        // value={
                        //   formState.imageLink && formState.imageLink.startsWith("data:")
                        //     ? formState.localImagePath
                        //     : formState.imageLink
                        // }
                        value={formState.imageLink}
                        placeholder={t("ImageURL")}
                        onChange={onImageLinkChanged}
                      />
                      <Button
                        style={{ gridColumn: "2", marginLeft: "5px" }}
                        onClick={handleUploadClick}
                        size="large"
                        appearance="secondary"
                        icon={<ArrowUpload24Regular />}
                      >
                        {t("Upload")}
                      </Button>
                      <input
                        type="file"
                        accept=".jpg, .jpeg, .png, .gif"
                        style={{ display: "none" }}
                        multiple={false}
                        onChange={handleImageSelection}
                        ref={fileInput}
                      />
                    </div>
                  </Field>
                  {/* <Text
                        className={formState.errorImageUrlMessage === "" ? "hide" : "show"}
                        error
                        size="small"
                        content={formState.errorImageUrlMessage}
                      /> */}
                  <Field size="large" label={t("Summary")}>
                    <Textarea
                      size="large"
                      appearance="filled-darker"
                      placeholder={t("Summary")}
                      value={formState.summary}
                      onChange={onSummaryChanged}
                    />
                  </Field>
                  <Field size="large" label={t("Author")}>
                    <Input
                      placeholder={t("Author")}
                      size="large"
                      onChange={onAuthorChanged}
                      autoComplete="off"
                      appearance="filled-darker"
                      value={formState.author}
                    />
                  </Field>
                  <Field size="large" label={t("ButtonTitle")}>
                    <Input
                      size="large"
                      placeholder={t("ButtonTitle")}
                      onChange={onBtnTitleChanged}
                      autoComplete="off"
                      appearance="filled-darker"
                      value={formState.buttonTitle}
                    />
                  </Field>
                  <Field size="large" label={t("ButtonURL")}>
                    <Input
                      size="large"
                      placeholder={t("ButtonURL")}
                      onChange={onBtnLinkChanged}
                      autoComplete="off"
                      appearance="filled-darker"
                      value={formState.buttonLink}
                    />
                  </Field>
                  {/* <Text
                        className={formState.errorButtonUrlMessage === "" ? "hide" : "show"}
                        error
                        size="small"
                        content={formState.errorButtonUrlMessage}
                      /> */}
                </div>
                <div className="card-area"></div>
              </div>
              <div className="footer-actions-inline">
                <div className="footer-action-right">
                  <Button
                    style={{ margin: "16px" }}
                    disabled={isNextBtnDisabled()}
                    id="saveBtn"
                    onClick={onNext}
                    appearance="primary"
                  >
                    {t("Next")}
                  </Button>
                </div>
              </div>
            </>
          )}
          {pageSelection === CurrentPageSelection.AudienceSelection && (
            <>
              <div className="adaptive-task-grid">
                <div className="form-area">
                  <Label id="labelId">
                    <h3>{t("SendHeadingText")}</h3>
                  </Label>
                  <RadioGroup aria-labelledby="labelId">
                    <Radio value={t("SendToGeneralChannel")} label={t("SendToGeneralChannel")} />
                    <div className={styles.root}>
                      <Label id={teamsComboId}>Pick Teams</Label>
                      {teamsSelectedOptions.length ? (
                        <ul id={teamsSelectedListId} className={styles.tagsList} ref={teamsSelectedListRef}>
                          {/* The "Remove" span is used for naming the buttons without affecting the Combobox name */}
                          <span id={`${teamsComboId}-remove`} hidden>
                            Remove
                          </span>
                          {teamsSelectedOptions.map((option, i) => (
                            <li key={option}>
                              <Button
                                size="small"
                                shape="circular"
                                appearance="primary"
                                icon={<Dismiss12Regular />}
                                iconPosition="after"
                                onClick={() => onTeamsTagClick(option, i)}
                                id={`${teamsComboId}-remove-${i}`}
                                aria-labelledby={`${teamsComboId}-remove ${teamsComboId}-remove-${i}`}
                              >
                                {option}
                              </Button>
                            </li>
                          ))}
                        </ul>
                      ) : null}
                      <Combobox
                        multiselect={true}
                        selectedOptions={teamsSelectedOptions}
                        appearance="filled-darker"
                        size="large"
                        onOptionSelect={onTeamsSelect}
                        ref={teamsComboboxInputRef}
                        aria-labelledby={teamsLabelledBy}
                        placeholder="Pick one or more teams"
                      >
                        {teams.map((opt) => (
                          <Option key={opt?.id}>{opt?.name}</Option>
                        ))}
                      </Combobox>
                    </div>
                    <Radio value={t("SendToRosters")} label={t("SendToRosters")} />
                    <div className={styles.root}>
                      <Label id={rostersComboId}>Pick Teams</Label>
                      {rostersSelectedOptions.length ? (
                        <ul id={rostersSelectedListId} className={styles.tagsList} ref={rostersSelectedListRef}>
                          {/* The "Remove" span is used for naming the buttons without affecting the Combobox name */}
                          <span id={`${rostersComboId}-remove`} hidden>
                            Remove
                          </span>
                          {rostersSelectedOptions.map((option, i) => (
                            <li key={option}>
                              <Button
                                size="small"
                                shape="circular"
                                appearance="primary"
                                icon={<Dismiss12Regular />}
                                iconPosition="after"
                                onClick={() => onRostersTagClick(option, i)}
                                id={`${rostersComboId}-remove-${i}`}
                                aria-labelledby={`${rostersComboId}-remove ${rostersComboId}-remove-${i}`}
                              >
                                {option}
                              </Button>
                            </li>
                          ))}
                        </ul>
                      ) : null}
                      <Combobox
                        multiselect={true}
                        selectedOptions={rostersSelectedOptions}
                        appearance="filled-darker"
                        size="large"
                        onOptionSelect={onRostersSelect}
                        ref={rostersComboboxInputRef}
                        aria-labelledby={rostersLabelledBy}
                        placeholder="Pick one or more teams"
                      >
                        {teams.map((opt) => (
                          <Option key={opt?.id}>{opt?.name}</Option>
                        ))}
                      </Combobox>
                    </div>
                    <Radio value={t("SendToAllUsers")} label={t("SendToAllUsers")} />
                    <Radio value={t("SendToGroups")} label={t("SendToGroups")} />
                    <div className={styles.root}>
                      <Label id={searchComboId}>Search Groups</Label>
                      {searchSelectedOptions.length ? (
                        <ul id={searchSelectedListId} className={styles.tagsList} ref={searchSelectedListRef}>
                          {/* The "Remove" span is used for naming the buttons without affecting the Combobox name */}
                          <span id={`${searchComboId}-remove`} hidden>
                            Remove
                          </span>
                          {searchSelectedOptions.map((option, i) => (
                            <li key={option}>
                              <Button
                                size="small"
                                shape="circular"
                                appearance="primary"
                                icon={<Dismiss12Regular />}
                                iconPosition="after"
                                onClick={() => onSearchTagClick(option, i)}
                                id={`${searchComboId}-remove-${i}`}
                                aria-labelledby={`${searchComboId}-remove ${searchComboId}-remove-${i}`}
                              >
                                {option}
                              </Button>
                            </li>
                          ))}
                        </ul>
                      ) : null}
                      <Combobox
                        appearance="filled-darker"
                        size="large"
                        onOptionSelect={onSearchSelect}
                        onChange={onSearchChange}
                        placeholder="Search for groups"
                      >
                        {queryGroups.map((opt) => (
                          <Option key={opt?.id}>{opt?.name}</Option>
                        ))}
                        {queryGroups.length === 0 && <Option disabled>No results</Option>}
                      </Combobox>
                    </div>
                  </RadioGroup>
                  {/* <RadioGroup
                        className="radioBtns"
                        checkedValue={formState.selectedRadioBtn}
                        onCheckedValueChange={onGroupSelected}
                        vertical={true}
                        items={[
                          {
                            name: "teams",
                            key: "teams",
                            value: "teams",
                            label: t("SendToGeneralChannel"),
                            children: (Component, { name, ...props }) => {
                              return (
                                <Flex key={name} column>
                                  <Component {...props} />
                                  <Dropdown
                                    hidden={!formState.teamsOptionSelected}
                                    placeholder={t("SendToGeneralChannelPlaceHolder")}
                                    search
                                    multiple
                                    items={getItems()}
                                    value={formState.selectedTeams}
                                    onChange={onTeamsChange}
                                    noResultsMessage={t("NoMatchMessage")}
                                  />
                                </Flex>
                              );
                            },
                          },
                          {
                            name: "rosters",
                            key: "rosters",
                            value: "rosters",
                            label: t("SendToRosters"),
                            children: (Component, { name, ...props }) => {
                              return (
                                <Flex key={name} column>
                                  <Component {...props} />
                                  <Dropdown
                                    hidden={!formState.rostersOptionSelected}
                                    placeholder={t("SendToRostersPlaceHolder")}
                                    search
                                    multiple
                                    items={getItems()}
                                    value={formState.selectedRosters}
                                    onChange={onRostersChange}
                                    unstable_pinned={formState.unstablePinned}
                                    noResultsMessage={t("NoMatchMessage")}
                                  />
                                </Flex>
                              );
                            },
                          },
                          {
                            name: "allUsers",
                            key: "allUsers",
                            value: "allUsers",
                            label: t("SendToAllUsers"),
                            children: (Component, { name, ...props }) => {
                              return (
                                <Flex key={name} column>
                                  <Component {...props} />
                                  <div className={formState.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                    <div className="noteText">
                                      <Text error content={t("SendToAllUsersNote")} />
                                    </div>
                                  </div>
                                </Flex>
                              );
                            },
                          },
                          {
                            name: "groups",
                            key: "groups",
                            value: "groups",
                            label: t("SendToGroups"),
                            children: (Component, { name, ...props }) => {
                              return (
                                <Flex key={name} column>
                                  <Component {...props} />
                                  <div
                                    className={formState.groupsOptionSelected && !formState.groupAccess ? "" : "hide"}
                                  >
                                    <div className="noteText">
                                      <Text error content={t("SendToGroupsPermissionNote")} />
                                    </div>
                                  </div>
                                  <Dropdown
                                    className="hideToggle"
                                    hidden={!formState.groupsOptionSelected || !formState.groupAccess}
                                    placeholder={t("SendToGroupsPlaceHolder")}
                                    search={onGroupSearch}
                                    multiple
                                    loading={formState.loading}
                                    loadingMessage={t("LoadingText")}
                                    items={getGroupItems()}
                                    value={formState.selectedGroups}
                                    onSearchQueryChange={onGroupSearchQueryChange}
                                    onChange={onGroupsChange}
                                    noResultsMessage={formState.noResultMessage}
                                    unstable_pinned={formState.unstablePinned}
                                  />
                                  <div
                                    className={formState.groupsOptionSelected && formState.groupAccess ? "" : "hide"}
                                  >
                                    <div className="noteText">
                                      <Text error content={t("SendToGroupsNote")} />
                                    </div>
                                  </div>
                                </Flex>
                              );
                            },
                          },
                        ]}
                      ></RadioGroup> */}
                </div>
                <div className="card-area"></div>
              </div>
              <div>
                <div className="footer-actions-inline">
                  <div className="footer-action-left">
                    <Button id="backBtn" onClick={onBack} appearance="secondary">
                      {t("Back")}
                    </Button>
                  </div>
                  <div className="footer-action-right">
                    <div className="footer-actions-flex">
                      <Spinner
                        id="draftingLoader"
                        size="small"
                        label={t("DraftingMessageLabel")}
                        labelPosition="after"
                      />
                      <Button
                        style={{ margin: "16px" }}
                        disabled={isSaveBtnDisabled()}
                        id="saveBtn"
                        onClick={onSave}
                        appearance="primary"
                      >
                        {t("SaveAsDraft")}
                      </Button>
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      )}
    </>
  );
};
