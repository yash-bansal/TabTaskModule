// import { useState } from "react";
import { Button, LinkIcon, Text } from "@fluentui/react-northstar";
import '../tabStyles.css';

import * as microsoftTeams from "@microsoft/teams-js";
import { UrlDialogInfo, dialog, DialogDimension, DialogInfo } from "@microsoft/teams-js";
import { useState } from "react";

const Welcome = (props: any) => {

  const _regItems: any[] = [];
  const _skillsItems: any[] = [];
  const initialState = {
    registrations: _regItems,
    skillsSubmissions: _skillsItems
  }
  const [state, setState] = useState(initialState);

  const baseUrl = `https://${window.location.hostname}${window.location.port !== '' && window.location.port !== null ? `:${window.location.port}` : ''}`;
  const invokeRegisterHTMLTask = () => {
    const urlDialogInfo: UrlDialogInfo = {
      title: 'Submit Employee Information',
      size: { height: DialogDimension.Medium, width: DialogDimension.Medium },
      url: baseUrl + "/index.html#/employeeInfoForm",
      fallbackUrl: ''
    };

    const submitHandler: dialog.DialogSubmitHandler = (dialogResponse: any) => {
      // console.log(`Submit handler - err: ${dialogResponse.err}`);
      // alert("Result = " + JSON.stringify(dialogResponse.result) + "\nError = " + JSON.stringify(dialogResponse.err));

      let items = state.registrations;
      items.push(dialogResponse.result);
      setState({ ...state, registrations: items });
    };

    microsoftTeams.dialog.open(urlDialogInfo, submitHandler);
  }

  const invokeACTask = () => {
    let taskInfo: DialogInfo = {
      title: "Submit Skills Information'",
      height: DialogDimension.Medium,
      width: DialogDimension.Medium,
      card: '',
      fallbackUrl: '',
      completionBotId: '',
    };

    taskInfo.card = `{
      "type": "AdaptiveCard",
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "version": "1.4",
      "body": [
          {
              "type": "Input.Text",
              "id": "name",
              "placeholder": "Enter full name",
              "label": "Name"
          },
          {
              "type": "Input.Text",
              "placeholder": "Enter role",
              "id": "role",
              "label": "Role"
          },
          {
              "type": "Input.ChoiceSet",
              "choices": [
                  {
                      "title": "SharePoint",
                      "value": "SharePoint"
                  },
                  {
                      "title": "SPFx",
                      "value": "SPFx"
                  },
                  {
                      "title": "Azure",
                      "value": "Azure"
                  },
                  {
                      "title": "DevOps",
                      "value": "DevOps"
                  }
              ],
              "placeholder": "-- Select --",
              "isMultiSelect": true,
              "id": "skills",
              "label": "Skills"
          },
          {
              "type": "ActionSet",
              "actions": [
                  {
                      "type": "Action.Submit",
                      "title": "Submit",
                      "data": {
                          "action": "submitskills"
                      },
                      "id": "submit"
                  }
              ],
              "spacing": "Large",
              "separator": true
          }
      ]
  }`;

    const submitHandler: ((err: string, result: string | object) => void) = (err, result) => {
      // console.log(`Submit handler - err: ${err}`);
      // alert("Result = " + JSON.stringify(result) + "\nError = " + JSON.stringify(err));

      let items = state.skillsSubmissions;
      items.push(result);
      setState({ ...state, skillsSubmissions: items });
    };

    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
  }

  const cards = [{
    title: "Introduction to Microsoft Teams Toolkit",
    url: "z8-cMEz350E"
  },
  {
    title: "Build your first Notification bot",
    url: "bwyd46tVzQo"
  },
  {
    title: "Getting started with Teams Toolkit for Visual Studio",
    url: "7oa0hW5pXt8"
  }
  ]
  const showVideo = (videoUrl: string) => {
    const urlDialogInfo: UrlDialogInfo = {
      title: 'Submit Employee Information',
      size: { height: DialogDimension.Large, width: DialogDimension.Large },
      url: baseUrl + "/videoPlayer.html?v=" + videoUrl,
      fallbackUrl: ''
    };
    microsoftTeams.dialog.open(urlDialogInfo);
  }

  return (
    <div className={`ms-Grid`}
      style={{ background: "radial-gradient(circle, rgba(238,174,202,1) 0%, rgba(148,187,233,1) 100%)" }}>
      <div className={`ms-Grid-row`} style={{ background: "linear-gradient(90deg, rgba(0,0,0,1) 30%, rgba(255,255,255,1) 70%, rgba(48,213,200,1) 100%)", padding: "2rem 5rem 3rem 5rem" }}>
        <div className={`ms-Grid-col ms-sm12`}>
          <Text style={{ color: "#30D5C8", fontSize: "4rem", fontWeight: "600", display: "block" }}>Wisdom</Text>
          <Text style={{ color: "white", fontSize: "1.25rem", display: "block" }}>Learn. Practice. Deliver.</Text>
        </div>
      </div>
      <div className={`ms-Grid-row`}>
        {
          cards.map((item) => {
            return <div className={`ms-Grid-col ms-sm3`}>
              <div className={`ms-Grid-row videoTile`}>
                <div className={`ms-Grid-col ms-sm12 videoTitle`}>
                  <Text>{item.title}</Text>
                </div>
                <div className={`ms-Grid-col ms-sm6`}>
                  <Button primary onClick={() => showVideo(item.url)} style={{ marginRight: "1rem" }}>View</Button>
                </div>
                <div className={`ms-Grid-col ms-sm6`} style={{ textAlign: "right" }}>
                  <LinkIcon onClick={() => {
                    var taskURL = baseUrl + "/videoPlayer.html?v=" + item.url;
                    const deepLink = `https://teams.microsoft.com/l/task/8e83acb9-ca8f-4642-88c0-a33e28a94e82?url=${taskURL}&height=1000&width=1000&title=YouTube Player`
                    const listener = (e: ClipboardEvent) => {
                      e.clipboardData?.setData('text/plain', deepLink);
                      e.preventDefault();
                      document.removeEventListener('copy', listener);
                    };
                    document.addEventListener('copy', listener);
                    document.execCommand('copy');
                    alert('Link Copied!')
                  }}
                    style={{ cursor: "pointer" }}
                  />
                </div>
              </div>
            </div>
          })
        }
      </div>
      <div className={`ms-Grid-row`} style={{ padding: "1rem" }}>
        <div className={`ms-Grid-col ms-sm2`}>
          <Button primary onClick={() => invokeRegisterHTMLTask()}>Register</Button>
        </div>
        <div className={`ms-Grid-col ms-sm10`}>
          {
            state.registrations.length > 0 ? (
              <div className={`ms-Grid-row`}>
                <div className={`ms-Grid-col ms-sm8`}>
                  <div className={`ms-Grid-row`} style={{ background: "#00000030", padding: "0.25rem" }}>
                    <div className={`ms-Grid-col ms-sm4`}>
                      <Text style={{ fontSize: "1rem" }}>Participant Name</Text>
                    </div>
                    <div className={`ms-Grid-col ms-sm4`}>
                      <Text style={{ fontSize: "1rem" }}>Participant Role</Text>
                    </div>
                    <div className={`ms-Grid-col ms-sm4`}>
                      <Text style={{ fontSize: "1rem" }}>Participant Topic</Text>
                    </div>
                  </div>
                  {
                    state.registrations.map((item) => {
                      return <div className={`ms-Grid-row`}
                        style={{ background: "#FFFFFF50", padding: "0.25rem", borderBottom: "2px solid #00000030" }}>
                        <div className={`ms-Grid-col ms-sm4`}>
                          <Text style={{ fontSize: "1rem" }}>{item.Name}</Text>
                        </div>
                        <div className={`ms-Grid-col ms-sm4`}>
                          <Text style={{ fontSize: "1rem" }}>{item.Role}</Text>
                        </div>
                        <div className={`ms-Grid-col ms-sm4`}>
                          <Text style={{ fontSize: "1rem" }}>{item.Topic}</Text>
                        </div>
                      </div>
                    })
                  }
                </div>
              </div>
            ) : (
              <div className={`ms-Grid-row`}>
                <div className={`ms-Grid-col ms-sm12`}>
                  <Text style={{ fontSize: "1rem" }}>No Registrations Found. Submit participant entries using Register button.</Text>
                </div>
              </div>
            )
          }
        </div>
      </div>
      <div className={`ms-Grid-row`} style={{ padding: "1rem 1rem 10rem 1rem" }}>
        <div className={`ms-Grid-col ms-sm2`}>
          <Button primary onClick={() => invokeACTask()}>Submit Skills</Button>
        </div>
        <div className={`ms-Grid-col ms-sm10`}>
          {
            state.skillsSubmissions.length > 0 ? (
              <div className={`ms-Grid-row`}>
                <div className={`ms-Grid-col ms-sm8`}>
                  <div className={`ms-Grid-row`} style={{ background: "#00000030", padding: "0.25rem" }}>
                    <div className={`ms-Grid-col ms-sm4`}>
                      <Text style={{ fontSize: "1rem" }}>Name</Text>
                    </div>
                    <div className={`ms-Grid-col ms-sm4`}>
                      <Text style={{ fontSize: "1rem" }}>Role</Text>
                    </div>
                    <div className={`ms-Grid-col ms-sm4`}>
                      <Text style={{ fontSize: "1rem" }}>Skills</Text>
                    </div>
                  </div>
                  {
                    state.skillsSubmissions.map((item) => {
                      return <div className={`ms-Grid-row`}
                        style={{ background: "#FFFFFF50", padding: "0.25rem", borderBottom: "2px solid #00000030" }}>
                        <div className={`ms-Grid-col ms-sm4`}>
                          <Text style={{ fontSize: "1rem" }}>{item.name}</Text>
                        </div>
                        <div className={`ms-Grid-col ms-sm4`}>
                          <Text style={{ fontSize: "1rem" }}>{item.role}</Text>
                        </div>
                        <div className={`ms-Grid-col ms-sm4`}>
                          <Text style={{ fontSize: "1rem" }}>{item.skills}</Text>
                        </div>
                      </div>
                    })
                  }
                </div>
              </div>
            ) : (
              <div className={`ms-Grid-row`}>
                <div className={`ms-Grid-col ms-sm12`}>
                  <Text style={{ fontSize: "1rem" }}>No Skills Submission Found. Submit entries using Submit Skills button.</Text>
                </div>
              </div>
            )
          }
        </div>
      </div>
    </div>
  );
}

export default Welcome;
