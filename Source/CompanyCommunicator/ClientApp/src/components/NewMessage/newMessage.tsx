// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./newMessage.scss";

import * as AdaptiveCards from "adaptivecards";
import * as React from "react";
import { useTranslation } from "react-i18next";
import { useParams } from "react-router-dom";
import {
  Spinner,
  Field,
  Input,
  Button,
  Textarea,
  RadioGroup,
  Radio,
  Label,
  Combobox,
  Option,
} from "@fluentui/react-components";
import * as microsoftTeams from "@microsoft/teams-js";
import { GetGroupsAction, GetTeamsDataAction, VerifyGroupAccessAction } from "../../actions";
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

import parse from "html-react-parser";

const validImageTypes = ["image/gif", "image/jpeg", "image/png", "image/jpg"];

type dropdownItem = {
  key: string;
  header: string;
  content: string;
  image: string;
  team: {
    id: string;
  };
};

export interface IDraftMessageState {
  id?: string;
  title: string;
  imageLink?: string;
  summary?: string;
  author: string;
  buttonTitle?: string;
  buttonLink?: string;
  teams: any[];
  rosters: any[];
  groups: any[];
  allUsers: boolean;
}

export interface formState {
  title: string;
  summary?: string;
  btnLink?: string;
  imageLink?: string;
  localImagePath?: string;
  btnTitle?: string;
  author: string;
  card?: any;
  page: string;
  teamsOptionSelected: boolean;
  rostersOptionSelected: boolean;
  allUsersOptionSelected: boolean;
  groupsOptionSelected: boolean;
  teams?: any[];
  groups?: any[];
  exists?: boolean;
  messageId: string;
  loader: boolean;
  loading: boolean;
  noResultMessage: string;
  unstablePinned?: boolean;
  selectedTeamsNum: number;
  selectedRostersNum: number;
  selectedGroupsNum: number;
  selectedRadioBtn: string;
  selectedTeams: dropdownItem[];
  selectedRosters: dropdownItem[];
  selectedGroups: dropdownItem[];
  errorImageUrlMessage: string;
  errorButtonUrlMessage: string;
}

let card: any;
// let renderCard: string = "<span></span>";

export const NewMessage = () => {
  let fileInput = React.createRef();
  const MAX_SELECTED_TEAMS_NUM: number = 20;
  const { t } = useTranslation();
  const { id } = useParams() as any;
  const teams = useAppSelector((state: RootState) => state.messages).teamsData.payload;
  const groups = useAppSelector((state: RootState) => state.messages).groups.payload;
  // const verifyGroupAccess = useAppSelector((state: RootState) => state.messages).verifyGroup.payload;
  const [renderCard, setRenderCard] = React.useState("<span></span>");

  const [loader, setLoader] = React.useState(false);

  const [formState, setFormState] = React.useState<formState>({
    title: "",
    summary: "",
    author: "",
    btnLink: "",
    imageLink: "",
    localImagePath: "",
    btnTitle: "",
    card: card,
    page: "CardCreation",
    teamsOptionSelected: true,
    rostersOptionSelected: false,
    allUsersOptionSelected: false,
    groupsOptionSelected: false,
    messageId: "",
    loader: true,
    loading: false,
    noResultMessage: "",
    unstablePinned: true,
    selectedTeamsNum: 0,
    selectedRostersNum: 0,
    selectedGroupsNum: 0,
    selectedRadioBtn: "teams",
    selectedTeams: [],
    selectedRosters: [],
    selectedGroups: [],
    errorImageUrlMessage: "",
    errorButtonUrlMessage: "",
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
      GetGroupsAction(dispatch, id);
      getDraftNotificationItem(id);
    }
  }, [id]);

  React.useEffect(() => {
    if (teams) {
      makeDropdownItemList(formState.selectedTeams, teams);
      makeDropdownItemList(formState.selectedRosters, teams);
    }
  }, [teams, formState.selectedTeams, formState.selectedRosters]);

  React.useEffect(() => {
    if (groups) {
      makeDropdownItems(groups);
    }
  }, [groups]);

  const getDraftNotificationItem = async (id: number) => {
    try {
      await getDraftNotification(id).then((response) => {
        const draftMessageDetail = response.data;
        let selectedRadioButton = "teams";
        if (draftMessageDetail.rosters.length > 0) {
          selectedRadioButton = "rosters";
        } else if (draftMessageDetail.groups.length > 0) {
          selectedRadioButton = "groups";
        } else if (draftMessageDetail.allUsers) {
          selectedRadioButton = "allUsers";
        }
        setFormState({
          ...formState,
          teamsOptionSelected: draftMessageDetail.teams.length > 0,
          selectedTeamsNum: draftMessageDetail.teams.length,
          rostersOptionSelected: draftMessageDetail.rosters.length > 0,
          selectedRostersNum: draftMessageDetail.rosters.length,
          groupsOptionSelected: draftMessageDetail.groups.length > 0,
          selectedGroupsNum: draftMessageDetail.groups.length,
          selectedRadioBtn: selectedRadioButton,
          selectedTeams: draftMessageDetail.teams,
          selectedRosters: draftMessageDetail.rosters,
          selectedGroups: draftMessageDetail.groups,
          title: draftMessageDetail.title,
          summary: draftMessageDetail.summary,
          btnLink: draftMessageDetail.buttonLink,
          imageLink: draftMessageDetail.imageLink,
          btnTitle: draftMessageDetail.buttonTitle,
          author: draftMessageDetail.author,
          allUsersOptionSelected: draftMessageDetail.allUsers,
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
    let adaptiveCard = new AdaptiveCards.AdaptiveCard();
    adaptiveCard.parse(card);
    if (formState.btnLink) {
      let link = formState.btnLink;
      adaptiveCard.onExecuteAction = function (action) {
        window.open(link, "_blank");
      };
    }
    setRenderCard(adaptiveCard.render().outerHTML);
  };

  const makeDropdownItems = (items: any[] | undefined) => {
    const resultedTeams: dropdownItem[] = [];
    if (items) {
      items.forEach((element) => {
        resultedTeams.push({
          key: element.id,
          header: element.name,
          content: element.mail,
          image: ImageUtil.makeInitialImage(element.name),
          team: {
            id: element.id,
          },
        });
      });
    }
    return resultedTeams;
  };

  const makeDropdownItemList = (items: any[], fromItems: any[] | undefined) => {
    const dropdownItemList: dropdownItem[] = [];
    items.forEach((element) =>
      dropdownItemList.push(
        typeof element !== "string"
          ? element
          : {
              key: fromItems!.find((x) => x.id === element).id,
              header: fromItems!.find((x) => x.id === element).name,
              image: ImageUtil.makeInitialImage(fromItems!.find((x) => x.id === element).name),
              team: {
                id: element,
              },
            }
      )
    );
    return dropdownItemList;
  };

  const getGroupItems = () => {
    if (formState.groups) {
      return makeDropdownItems(formState.groups);
    }
    const dropdownItems: dropdownItem[] = [];
    return dropdownItems;
  };

  const handleUploadClick = (event: any) => {
    if (fileInput.current) {
      // fileInput.current.click();
    }
  };

  const checkValidSizeOfImage = (resizedImageAsBase64: string) => {
    var stringLength = resizedImageAsBase64.length - "data:image/png;base64,".length;
    var sizeInBytes = 4 * Math.ceil(stringLength / 3) * 0.5624896334383812;
    var sizeInKb = sizeInBytes / 1000;

    if (sizeInKb <= 1024) return true;
    else return false;
  };

  // const handleImageSelection = () => {
  //   const file = fileInput.current.files[0];

  //   if (file) {
  //     const fileType = file["type"];
  //     const { type: mimeType } = file;

  //     if (!validImageTypes.includes(fileType)) {
  //       setFormState({ ...formState, errorImageUrlMessage: t("ErrorImageTypesMessage") });
  //       return;
  //     }

  //     setFormState({ ...formState, localImagePath: file["name"] });
  //     setFormState({ ...formState, errorImageUrlMessage: "" });

  //     const fileReader = new FileReader();
  //     fileReader.readAsDataURL(file);
  //     fileReader.onload = () => {
  //       var image = new Image();
  //       image.src = fileReader.result as string;
  //       var resizedImageAsBase64 = fileReader.result as string;

  //       image.onload = function (e: any) {
  //         const MAX_WIDTH = 1024;

  //         if (image.width > MAX_WIDTH) {
  //           const canvas = document.createElement("canvas");
  //           canvas.width = MAX_WIDTH;
  //           canvas.height = ~~(image.height * (MAX_WIDTH / image.width));
  //           const context = canvas.getContext("2d", { alpha: false });
  //           if (!context) {
  //             return;
  //           }
  //           context.drawImage(image, 0, 0, canvas.width, canvas.height);
  //           resizedImageAsBase64 = canvas.toDataURL(mimeType);
  //         }
  //       };

  //       if (!checkValidSizeOfImage(resizedImageAsBase64)) {
  //         setFormState({ ...formState, errorImageUrlMessage: t("ErrorImageSizeMessage") });
  //         return;
  //       }

  //       setCardImageLink(card, resizedImageAsBase64);
  //       updateCard();
  //       setFormState({ ...formState, imageLink: resizedImageAsBase64 });
  //     };

  //     fileReader.onerror = (error) => {
  //       //reject(error);
  //     };
  //   }
  // };

  const onGroupSelected = (event: any, data: any) => {
    setFormState({
      ...formState,
      selectedRadioBtn: data.value,
      teamsOptionSelected: data.value === "teams",
      rostersOptionSelected: data.value === "rosters",
      groupsOptionSelected: data.value === "groups",
      allUsersOptionSelected: data.value === "allUsers",
      selectedTeams: data.value === "teams" ? formState.selectedTeams : [],
      selectedTeamsNum: data.value === "teams" ? formState.selectedTeamsNum : 0,
      selectedRosters: data.value === "rosters" ? formState.selectedRosters : [],
      selectedRostersNum: data.value === "rosters" ? formState.selectedRostersNum : 0,
      selectedGroups: data.value === "groups" ? formState.selectedGroups : [],
      selectedGroupsNum: data.value === "groups" ? formState.selectedGroupsNum : 0,
    });
  };

  const isSaveBtnDisabled = () => {
    const teamsSelectionIsValid =
      (formState.teamsOptionSelected && formState.selectedTeamsNum !== 0) || !formState.teamsOptionSelected;
    const rostersSelectionIsValid =
      (formState.rostersOptionSelected && formState.selectedRostersNum !== 0) || !formState.rostersOptionSelected;
    const groupsSelectionIsValid =
      (formState.groupsOptionSelected && formState.selectedGroupsNum !== 0) || !formState.groupsOptionSelected;
    const nothingSelected =
      !formState.teamsOptionSelected &&
      !formState.rostersOptionSelected &&
      !formState.groupsOptionSelected &&
      !formState.allUsersOptionSelected;
    return !teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected;
  };

  const isNextBtnDisabled = () => {
    const title = formState.title;
    const btnTitle = formState.btnTitle;
    const btnLink = formState.btnLink;
    return !(
      title &&
      ((btnTitle && btnLink) || (!btnTitle && !btnLink)) &&
      formState.errorImageUrlMessage === "" &&
      formState.errorButtonUrlMessage === ""
    );
  };

  const getItems = () => {
    const resultedTeams: dropdownItem[] = [];
    if (formState.teams) {
      let remainingUserTeams = formState.teams;
      if (formState.selectedRadioBtn !== "allUsers") {
        if (formState.selectedRadioBtn === "teams") {
          formState.teams.filter((x) => formState.selectedTeams.findIndex((y) => y.team.id === x.id) < 0);
        } else if (formState.selectedRadioBtn === "rosters") {
          formState.teams.filter((x) => formState.selectedRosters.findIndex((y) => y.team.id === x.id) < 0);
        }
      }
      remainingUserTeams.forEach((element) => {
        resultedTeams.push({
          key: element.id,
          header: element.name,
          content: element.mail,
          image: ImageUtil.makeInitialImage(element.name),
          team: {
            id: element.id,
          },
        });
      });
    }
    return resultedTeams;
  };

  const onTeamsChange = (event: any, itemsData: any) => {
    if (itemsData.value.length > MAX_SELECTED_TEAMS_NUM) return;
    setFormState({
      ...formState,
      selectedTeams: itemsData.value,
      selectedTeamsNum: itemsData.value.length,
      selectedRosters: [],
      selectedRostersNum: 0,
      selectedGroups: [],
      selectedGroupsNum: 0,
    });
  };

  const onRostersChange = (event: any, itemsData: any) => {
    if (itemsData.value.length > MAX_SELECTED_TEAMS_NUM) return;
    setFormState({
      ...formState,
      selectedRosters: itemsData.value,
      selectedRostersNum: itemsData.value.length,
      selectedTeams: [],
      selectedTeamsNum: 0,
      selectedGroups: [],
      selectedGroupsNum: 0,
    });
  };

  const onGroupsChange = (event: any, itemsData: any) => {
    setFormState({
      ...formState,
      selectedGroups: itemsData.value,
      selectedGroupsNum: itemsData.value.length,
      groups: [],
      selectedTeams: [],
      selectedTeamsNum: 0,
      selectedRosters: [],
      selectedRostersNum: 0,
    });
  };

  const onGroupSearch = (itemList: any, searchQuery: string) => {
    const result = itemList.filter(
      (item: { header: string; content: string }) =>
        (item.header && item.header.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1) ||
        (item.content && item.content.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1)
    );
    return result;
  };

  const onGroupSearchQueryChange = async (event: any, itemsData: any) => {
    if (!itemsData.searchQuery) {
      setFormState({ ...formState, groups: [], noResultMessage: "" });
    } else if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
      setFormState({ ...formState, loading: false, noResultMessage: t("NoMatchMessage") });
    } else if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
      // handle event trigger on item select.
      const result =
        itemsData.items &&
        itemsData.items.find(
          (item: { header: string }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
        );
      if (result) {
        return;
      }

      setFormState({ ...formState, loading: true, noResultMessage: "" });

      try {
        const query = encodeURIComponent(itemsData.searchQuery);
        const response = await searchGroups(query);
        setFormState({ ...formState, groups: response.data, loading: false, noResultMessage: t("NoMatchMessage") });
      } catch (error) {
        return error;
      }
    }
  };

  const onSave = () => {
    const selectedTeams: string[] = [];
    const selctedRosters: string[] = [];
    const selectedGroups: string[] = [];
    formState.selectedTeams.forEach((x) => selectedTeams.push(x.team.id));
    formState.selectedRosters.forEach((x) => selctedRosters.push(x.team.id));
    formState.selectedGroups.forEach((x) => selectedGroups.push(x.team.id));

    const draftMessage: IDraftMessageState = {
      id: formState.messageId,
      title: formState.title,
      imageLink: formState.imageLink,
      summary: formState.summary,
      author: formState.author,
      buttonTitle: formState.btnTitle,
      buttonLink: formState.btnLink,
      teams: selectedTeams,
      rosters: selctedRosters,
      groups: selectedGroups,
      allUsers: formState.allUsersOptionSelected,
    };

    let spanner = document.getElementsByClassName("draftingLoader");
    spanner[0].classList.remove("hiddenLoader");

    if (formState.exists) {
      editDraftMessage(draftMessage).then(() => {
        microsoftTeams.tasks.submitTask();
      });
    } else {
      postDraftMessage(draftMessage).then(() => {
        microsoftTeams.tasks.submitTask();
      });
    }
  };

  const editDraftMessage = async (draftMessage: IDraftMessageState) => {
    try {
      await updateDraftNotification(draftMessage);
    } catch (error) {
      return error;
    }
  };

  const postDraftMessage = async (draftMessage: IDraftMessageState) => {
    try {
      await createDraftNotification(draftMessage);
    } catch (error) {
      throw error;
    }
  };

  const onNext = (event: any) => {
    setFormState({ ...formState, page: "AudienceSelection" });

    // updateCard();
  };

  const onBack = (event: any) => {
    setFormState({ ...formState, page: "CardCreation" });
    // updateCard();
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
    setCardBtn(card, event.target.value, formState.btnLink);
    setFormState({ ...formState, btnTitle: event.target.value });
    updateAdaptiveCard();
  };

  const onBtnLinkChanged = (event: any) => {
    setCardBtn(card, formState.btnTitle, event.target.value);
    setFormState({ ...formState, btnLink: event.target.value });
    updateAdaptiveCard();
  };

  return (
    <>
      {loader && <Spinner labelPosition="below" size="large" />}
      {!loader && (
        <div>
          {formState.page === "CardCreation" && (
            <>
              <div className="adaptive-task-grid">
                <div className="form-area">
                  <Field label={t("TitleText")}>
                    <Input
                      placeholder={t("PlaceHolderTitle")}
                      onChange={onTitleChanged}
                      autoComplete="off"
                      value={formState.title}
                    />
                  </Field>
                  <Field label={t("ImageURL")}>
                    <input
                      className="file-button"
                      aria-labelledby="imageLabelId"
                      type="file"
                      id="file"
                      name="file"
                      aria-label={t("ImageURL")}
                    />
                    {/* <Input
                          fluid
                          className="inputField imageField"
                          value={
                            formState.imageLink && formState.imageLink.startsWith("data:")
                              ? formState.localImagePath
                              : formState.imageLink
                          }
                          label={
                            <>
                              {t("ImageURL")}
                              <TooltipHost
                                content={t("ImageSizeInfoContent")}
                                calloutProps={{ gapSpace: 0 }}
                                hostClassName="tooltipHostStyles"
                              >
                                <Icon aria-label="Info" iconName="Info" className="tooltipHostStylesInsideContent" />
                              </TooltipHost>
                            </>
                          }
                          placeholder={t("ImageURL")}
                          onChange={onImageLinkChanged}
                          error={!(formState.errorImageUrlMessage === "")}
                          autoComplete="off"
                        />

                        <Flex.Item push>
                          <Button
                            onClick={handleUploadClick}
                            size="medium"
                            className="inputField"
                            content={t("Upload")}
                            iconPosition="before"
                          />
                        </Flex.Item>
                        <input
                          type="file"
                          accept=".jpg, .jpeg, .png, .gif"
                          style={{ display: "none" }}
                          multiple={false}
                          onChange={handleImageSelection}
                          ref={fileInput}
                        /> */}
                  </Field>
                  {/* <Text
                        className={formState.errorImageUrlMessage === "" ? "hide" : "show"}
                        error
                        size="small"
                        content={formState.errorImageUrlMessage}
                      /> */}
                  <Field label={t("Summary")}>
                    <Textarea placeholder={t("Summary")} value={formState.summary} onChange={onSummaryChanged} />
                  </Field>
                  <Field label={t("Author")}>
                    <Input
                      placeholder={t("Author")}
                      onChange={onAuthorChanged}
                      autoComplete="off"
                      value={formState.author}
                    />
                  </Field>
                  <Field label={t("ButtonTitle")}>
                    <Input
                      placeholder={t("ButtonTitle")}
                      onChange={onBtnTitleChanged}
                      autoComplete="off"
                      value={formState.btnTitle}
                    />
                  </Field>
                  <Field label={t("ButtonURL")}>
                    <Input
                      placeholder={t("ButtonURL")}
                      onChange={onBtnLinkChanged}
                      autoComplete="off"
                      value={formState.btnLink}
                    />
                  </Field>
                  {/* <Text
                        className={formState.errorButtonUrlMessage === "" ? "hide" : "show"}
                        error
                        size="small"
                        content={formState.errorButtonUrlMessage}
                      /> */}
                </div>
                <div className="card-area">{parse(renderCard || "<span></span>")}</div>
              </div>
              <div>
                <Button
                  style={{ float: "right" }}
                  disabled={isNextBtnDisabled()}
                  id="saveBtn"
                  onClick={onNext}
                  appearance="primary"
                >
                  {t("Next")}
                </Button>
              </div>
            </>
          )}
          {formState.page === "AudienceSelection" && (
            <>
              <div className="adaptive-task-grid">
                <div className="form-area">
                  <Label id="labelId">
                    <h3>{t("SendHeadingText")}</h3>
                  </Label>
                  <RadioGroup aria-labelledby="labelId">
                    <Radio value={t("SendToGeneralChannel")} label={t("SendToGeneralChannel")} />
                    <Label id={"comboId"}>Best pet</Label>
                    <Combobox aria-labelledby={"comboId"} placeholder="Select an animal">
                      {getItems().map((opt) => (
                        <Option key={opt.key}>{opt.header}</Option>
                      ))}
                    </Combobox>
                    <Radio value={t("SendToRosters")} label={t("SendToRosters")} />
                    <Radio value={t("SendToAllUsers")} label={t("SendToAllUsers")} />
                    <Radio value={t("SendToGroups")} label={t("SendToGroups")} />
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
                <div className="card-area">{parse(renderCard || "<span></span>")}</div>
              </div>
              <Spinner
                id="draftingLoader"
                className="hiddenLoader draftingLoader"
                size="small"
                label={t("DraftingMessageLabel")}
                labelPosition="after"
              />
              <Button id="backBtn" onClick={onBack} appearance="secondary">
                {t("Back")}
              </Button>
              <Button disabled={isSaveBtnDisabled()} id="saveBtn" onClick={onSave} appearance="primary">
                {t("SaveAsDraft")}
              </Button>
            </>
          )}
          {formState.page !== "CardCreation" && formState.page !== "AudienceSelection" && <div>Error</div>}
        </div>
      )}
    </>
  );
};
