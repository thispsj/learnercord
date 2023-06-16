import {
    TeamsActivityHandler,
    TurnContext,
    MessageFactory,
    CardFactory, MessagingExtensionAction, MessagingExtensionActionResponse, MessagingExtensionAttachment, MessagingExtensionResponse
  } from "botbuilder";
  
  import * as Util from "util";
  import * as debug from "debug";
  import { v4 as uuidgen } from "uuid";
  import axios from "axios";
  
  const TextEncoder = Util.TextEncoder;
  const log = debug("msteams");
  const msGraphApiKey = process.env.MS_GRAPH_API_KEY;
  const msGraphAppId = process.env.MS_GRAPH_CLIENT_ID;
  const msGraphDirectoryId = process.env.MS_GRAPH_DIRECTORY_ID;
  
  export class RABot extends TeamsActivityHandler {
    
    protected async handleTeamsMessagingExtensionFetchTask (context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
        
        const uuid = uuidgen();
        //Create a new excel file on OneDrive.
        const fileName = `Attendance-${uuid}.xlsx`;
        //Get access token to access OneDrive
        const accessToken = await this.getAccessToken();
        //Get user id of the teacher who initiated the task module.
        const userId = context.activity.from.id;
        const url = `https://graph.microsoft.com/v1.0/${userId}/drive/root/children/${fileName}/content`;
        const headers = {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };
        const response = await axios.put(url, null, { headers });
        const adaptiveCardSource = require("./cards/recordattendance.json");
        let aCJson = JSON.parse(adaptiveCardSource);
        aCJson.body[0].columns[0].items[2].actions[0].data.id=uuid;
        aCJson.body[0].columns[0].items[2].actions[0].data.user = userId;
        const adaptiveCard = CardFactory.adaptiveCard(adaptiveCardSource);

        const mresponse: MessagingExtensionActionResponse = {
            task: {
                type: "continue",
                value: {
                    card: adaptiveCard,
                    title: "Record Attendance",
                    height: 150,
                    width: 500
                }}} as MessagingExtensionResponse

                return Promise.resolve(mresponse);

    }

    private async getAccessToken(): Promise<string> {
        const url = `https://login.microsoftonline.com/${msGraphDirectoryId}/oauth2/v2.0/token`;
        const data = {
            client_id: msGraphAppId,
            scope: "https://graph.microsoft.com/.default",
            client_secret: msGraphApiKey,
            grant_type: "client_credentials"
        };
        const response = await axios.post(url, data);
        return response.data.access_token;
    }

    async handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
        if (context.activity.value && context.activity.value.attendanceMarked) {
          const userName = context.activity.from.name;
          // Use the Microsoft Graph API to record the attendance in an Excel file
            // on OneDrive.
            const accessToken = await this.getAccessToken();
            const userId = context.activity.from.id;
            const url = `https://graph.microsoft.com/v1.0/${userId}/drive/root/children`;
            const headers = {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json"
            }
            const response = await axios.get(url, { headers });
            const files = response.data.value;
            const attendanceFile = files.find((file) => file.name.startsWith("Attendance"));
          await context.sendActivity(`Attendance marked for ${userName}`);
        }
      }
  }