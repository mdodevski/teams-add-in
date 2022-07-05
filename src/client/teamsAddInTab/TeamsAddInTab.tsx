import * as React from "react";
import { Box, Button, Loader, Provider } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { HttpClient } from "../services/httpService";
import validator from "validator";
import { TextField } from "office-ui-fabric-react";
import { SendFileToSignMultiple } from "../models/SendFileToSignMultiple";

export const TeamsAddInTab = () => {
  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();
  const [user, setUser] = useState<any>();
  const [level, setLevel] = useState<any>();
  const [prevLevel, setPrevLevel] = useState<any>();
  const [root, setRoot] = useState<any>();
  const [file, setFile] = useState<any>();
  const [addressList, setAddressList] = useState<any>([{ Email: "" }]);

  const httpClient = new HttpClient();

  // handle input change
  const handleInputChange = (e, index) => {
    const { name, value } = e.target;
    const list = [...addressList];
    list[index][name] = value.toLowerCase();
    setAddressList(list);
  };

  const handleRemoveClick = (index) => {
    const list = [...addressList];
    list.splice(index, 1);
    setAddressList(list);
  };

  // handle click event of the Add button
  const handleAddClick = () => {
    setAddressList([...addressList, { Email: "" }]);
  };

  const [emailsValid, setEmailValidation] = useState<any>(null);

  const validateEmail = (e, index) => {
    handleInputChange(e, index);
    setEmailValidation(true);
    addressList.forEach((email) => {
      if (!validator.isEmail(email.Email)) {
        setEmailValidation(false);
      }
    });
  };

  const setResult = (res: any) => {
    httpClient.GetUser(res).then((res2) => setUser(res2));

    res.GroupId = context?.groupId;
    res.ChannelId = context?.channelId;

    httpClient.GetChannelRootDirectory(res).then((res2) => {
      setRoot(res2);

      res.DriveId = res2.driveId;
      res.ItemId = res2.id;

      httpClient.GetChannelFolderChildren(res).then((res3) => {
        setPrevLevel(res3);
        setLevel(res3);
      });
    });
  };

  function getLevel(res: any) {
    let request = {
      assertion: res.assertion,
      client_id: res.client_id,
      client_secret: res.client_secret,
      grant_type: res.grant_type,
      requested_token_use: res.requested_token_use,
      scope: res.scope,
      DriveId: res.params.driveId,
      ItemId: res.params.id,
      Name: res.params.name,
      FileType: res.params.fileType,
    };

    if (res.params.type == "File") {
      httpClient.GetFileContent(request).then((res2) => {
        setFile(res2);
        console.log(res2)
      });
    } else if (res.params.type == "Folder") {
      if (res.params.childCount <= 0) {
        alert("The folder is empty!");
      } else {
        httpClient.GetChannelFolderChildren(request).then((res3) => {
          if (res3.length < 1) {
            alert("The folder is empty!");
          } else {
            setLevel(res3);
          }
        });
      }
    } else {
      alert("Error!");
    }
  }

  function prepareFileForSending(signItYourself: boolean) {
    var data: SendFileToSignMultiple = {
      Document: "",
      DocumentName: "",
      Signers: undefined,
      Initiator: undefined,
    };
    if (!signItYourself) {
      data = {
        Document: file.base64,
        DocumentName: file.name,
        Signers: addressList,
        Initiator: { Email: context?.userPrincipalName?.toLowerCase() },
      };
    } else {
      data = {
        Document: file.base64,
        DocumentName: file.name,
        Signers: [{ Email: context?.userPrincipalName?.toLowerCase() }],
        Initiator: { Email: context?.userPrincipalName?.toLowerCase() },
      };
    }

    httpClient.SendFileToSign(data).then((responseData) => {

      console.log(data);
      console.log(responseData);
      if (responseData == null || responseData == undefined) {
        setLevel(prevLevel);
        setFile(null);
      } else {
        if (data.Signers[0].Email === data.Initiator.Email) {
          window.open(responseData.Url, "_blank");
        } else {
          alert("Success!");
          setFile(null);
          setAddressList([{ Email: "" }]);
        }
      }
    });
  }

  useEffect(() => {
    if (inTeams === true) {
      microsoftTeams.getContext((context) => {});
      httpClient.GetAuthToken(setResult);

      microsoftTeams.appInitialization.notifySuccess();
    } else {
      setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

  useEffect(() => {
    if (context) {
      setEntityId(context.entityId);
    }
  }, [context]);

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <Provider theme={theme}>
      {!file &&(
        <div className="filesContainer">
          {root && (
            <div className="titleContainer">
              <span className="filesTitle">Files in {root.name} channel</span>{" "}
              {<Button onClick={() => setLevel(prevLevel)}>Go back</Button>}
            </div>
          )}
          {level &&
            level.map((item) => {
              return (
                <div
                  className="fileItem"
                  onClick={() => httpClient.GetAuthToken(getLevel, item)}
                >
                  <img
                    className="fileImage"
                    src={
                      item.fileType == "pdf"
                        ? "https://mdodevski-front.vizibit.eu/assets/pdf.png"
                        : item.fileType == "docx" || item.fileType == "doc"
                        ? "https://mdodevski-front.vizibit.eu/assets/doc.png"
                        : item.fileType == "ppt" || item.fileType == "odp"
                        ? "https://mdodevski-front.vizibit.eu/assets/ppt.png"
                        : item.fileType == "xls"
                        ? "https://mdodevski-front.vizibit.eu/assets/xls.png"
                        : item.type == "Folder"
                        ? "https://mdodevski-front.vizibit.eu/assets/folder.png"
                        : "https://mdodevski-front.vizibit.eu/assets/file.png"
                    }
                    alt="item_type"
                  />{" "}
                  <span className="fileName">{item.name}</span>
                  <span className="fileChildren">
                    {item.type == "Folder"
                      ? `${item.childCount} items inside`
                      : ""}
                  </span>
                </div>
              );
            })}
          {!level && <Loader></Loader>}
        </div>
      )}
      {file && (
        <div>
          <div className="titleItem">
            <span className="filesTitle">Chosen file:</span>
            <img
              className="fileImage"
              src={
                file.fileType == "pdf"
                  ? "https://mdodevski-front.vizibit.eu/assets/pdf.png"
                  : file.fileType == "docx" || file.fileType == "doc"
                  ? "https://mdodevski-front.vizibit.eu/assets/doc.png"
                  : file.fileType == "ppt" || file.fileType == "odp"
                  ? "https://mdodevski-front.vizibit.eu/assets/ppt.png"
                  : file.fileType == "xls"
                  ? "https://mdodevski-front.vizibit.eu/assets/xls.png"
                  : "https://mdodevski-front.vizibit.eu/assets/file.png"
              }
              alt="item_type"
            />{" "}
            <span className="fileName">{file.name}</span>
            {
              <Button
                onClick={() => {
                  setLevel(prevLevel), setFile(null);
                }}
              >
                Choose again
              </Button>
            }
          </div>

          <div className="signContainer">
            <span className="filesTitle">Sign it yourself</span>
            <Button onClick={() => prepareFileForSending(true)}>Sign</Button>
          </div>

          <div className="sendContainer">
            <span className="filesTitle">Send the file to be signed</span>
            {addressList.map((x, i) => {
              return (
                <div className="box">
                  <TextField
                    className="textBox"
                    id="address"
                    name="Email"
                    label="Enter email"
                    value={x.Email}
                    onChange={(e) => validateEmail(e, i)}
                  />

                  <div className="btn-box">
                    {addressList.length !== 1 && (
                      <Button
                        className="removeButton"
                        onClick={() => handleRemoveClick(i)}
                      >
                        Remove
                      </Button>
                    )}
                    {addressList.length - 1 === i && (
                      <Button onClick={handleAddClick}>Add</Button>
                    )}
                  </div>
                </div>
              );
            })}
            <div className="buttonContainer">
              <Button
                disabled={!emailsValid}
                className="sendButton"
                onClick={() => prepareFileForSending(false)}
              >
                Send
              </Button>
            </div>
            {!emailsValid && emailsValid != null && (
              <span className="errorTitle">Invalid email address!</span>
            )}
          </div>
        </div>
      )}
    </Provider>
  );
};
