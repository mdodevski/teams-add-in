import axios, { AxiosResponse } from "axios";
import * as microsoftTeams from "@microsoft/teams-js";
import { SendFileToSignMultiple } from "../models/SendFileToSignMultiple";

export class HttpClient {
  baseUrlMyApi = "https://mdodevski.vizibit.eu";
  baseUrlSignumId = "https://test.signumid.hr";
  baseUrlGraphApi = "https://graph.microsoft.com/v1.0";
  baseUrlMicrosoftOnline = "https://login.microsoftonline.com";

  async GetUser(request: any) {
    const res = await axios.post<any, AxiosResponse<any>>(
      this.baseUrlMyApi + "/v/1/signumid_integrations/get_user",
      request
    );
    return res.data;
  }

  async GetChannelRootDirectory(request: any) {
    const res = await axios.post<any, AxiosResponse<any>>(
      this.baseUrlMyApi + "/v/1/signumid_integrations/get_channel_root_folder",
      request
    );
    return res.data;
  }

  async GetChannelFolderChildren(request: any) {
    const res = await axios.post<any, AxiosResponse<any>>(
      this.baseUrlMyApi + "/v/1/signumid_integrations/get_channel_folder_children",
      request
    );
    return res.data;
  }

  async GetFileContent(request: any) {
    const res = await axios.post<any, AxiosResponse<any>>(
      this.baseUrlMyApi + "/v/1/signumid_integrations/get_file_content",
      request
    );
    return res.data;
  }

  async GetAuthToken(fn: Function, params?: any) {

    microsoftTeams.authentication.getAuthToken({
      successCallback: (result): void => {
        fn({
          client_id: "bbb71de5-d64e-4ad1-9994-40d0ff295dbb",
          client_secret: "2nR8Q~gfQhgQRyLvTMkO0XeWFisbdnCA.w4UTaly",
          scope:
            "user.read email openid profile Files.Read.All Files.ReadWrite.All Group.Read.All Group.ReadWrite.All offline_access",
          grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
          requested_token_use: "on_behalf_of",
          assertion: result,
          params
        });
      },
      failureCallback(error) {
        console.log(error);
      },
    });
  }

  async SendFileToSign(request: SendFileToSignMultiple) {
    console.log(request);
    const res = await axios
      .post<SendFileToSignMultiple, AxiosResponse<any>>(this.baseUrlSignumId + "/v/1/signature/workflow/pdf/sequential", request);
    return res.data;
  }
}
