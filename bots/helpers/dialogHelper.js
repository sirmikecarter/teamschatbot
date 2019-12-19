// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {CardFactory} = require('botbuilder');


class DialogHelper {

     createRAW1Purpose() {
       return CardFactory.adaptiveCard({
          "type": "AdaptiveCard",
          "body": [
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "Image",
                                  "style": "Person",
                                  "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                  "size": "Small"
                              }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "weight": "Bolder",
                                  "text": "What's the Purpose of this Request?",
                                  "wrap": true
                              }
                          ],
                          "width": "stretch"
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.ShowCard",
                  "title": "Architecture Approval",
                  "card": {
                      "type": "AdaptiveCard",
                      "body": [
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": 2,
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Select an Option",
                                              "isSubtle": true,
                                              "wrap": true
                                          }
                                      ]
                                  }
                              ]
                          }
                      ],
                      "actions": [
                          {
                              "type": "Action.Submit",
                              "id": "New",
                              "title": "New (i.e. Application Approval, Custom Solution, Policy)",
                              "data":{
                                    "action": "createRAW2TypeArch",
                                    "option": "New"
                              }
                          },
                          {
                              "type": "Action.Submit",
                              "id": "Change",
                              "title": "Change (i.e. Application Version Upgrade)",
                              "data":{
                                    "action": "createRAW2TypeArch",
                                    "option": "Change"
                              }
                          }
                      ],
                      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json"
                  }
              },
              {
                  "type": "Action.ShowCard",
                  "title": "Market Analysis",
                  "card": {
                      "type": "AdaptiveCard",
                      "body": [
                        {
                         "type": "TextBlock",
                         "text": 'What Type of Request Is This?',
                         "weight": "bolder",
                         "isSubtle": false
                        },
                        {
                             "type": "Input.ChoiceSet",
                             "id": "selectedValues",
                             "isMultiSelect": true,
                             "value": "1",
                             "choices": [
                                 {
                                     "title": "Market Research (i.e. Web Research, White Papers)",
                                     "value": "Market Research"
                                 },
                                 {
                                     "title": "Vendor Analysis (i.e. Top Vendors, Magic Quadrant)",
                                     "value": "Vendor Analysis"
                                 },
                                 {
                                     "title": "Schedule Vendor Demos",
                                     "value": "Schedule Vendor Demos"
                                 },
                                 {
                                     "title": "Build Prototype (i.e. RnD environment)",
                                     "value": "Build Prototype"
                                 }
                             ]
                         }
                      ],
                      "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Submit",
                            "data": {
                              "action": "createRAW2TypeMarket"
                            }
                        }
                      ],
                      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json"
                  }
              }
          ],
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "version": "1.0"
      });
     }

     createRAWVersionUpgrade() {
       return CardFactory.adaptiveCard({
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.0",
        "body": [
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "items": [
                          {
                              "type": "Image",
                              "style": "Person",
                              "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                              "size": "Small"
                          }
                      ],
                      "width": "auto"
                  },
                  {
                      "type": "Column",
                      "items": [
                          {
                              "type": "TextBlock",
                              "weight": "Bolder",
                              "text": "Please fill out the following",
                              "wrap": true
                          }
                      ],
                      "width": "stretch"
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": 2,
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "What's the Application Name?",
                              "isSubtle": false,
                              "wrap": true
                          },
                          {
                              "type": "Input.Text",
                              "id": "applicationName",
                              "placeholder": "Application Name"
                          },
                          {
                              "type": "TextBlock",
                              "text": "What's the Current Version of the Application?",
                              "isSubtle": false,
                              "wrap": true
                          },
                          {
                              "type": "Input.Text",
                              "id": "currentVersion",
                              "placeholder": "Current Version"
                          },
                          {
                              "type": "TextBlock",
                              "text": "What's the Proposed Version of the Application?",
                              "isSubtle": false,
                              "wrap": true
                          },
                          {
                              "type": "Input.Text",
                              "id": "proposedVersion",
                              "placeholder": "Proposed Version"
                          },
                          {
                              "type": "TextBlock",
                              "text": "Are there any new or changes to the existing solution?",
                              "isSubtle": false,
                              "wrap": true
                          }
                      ]
                  }
              ]
          },
          {
              "type": "Container",
              "items": [
                  {
                      "type": "ColumnSet",
                      "columns": [
                          {
                              "type": "Column",
                              "width": "stretch",
                              "items": [
                                  {
                                      "type": "TextBlock",
                                      "text": "Business or IT processes",
                                      "wrap": true,
                                      "id": "Text1"
                                  }
                              ]
                          },
                          {
                              "type": "Column",
                              "width": "auto",
                              "items": [
                                  {
                                      "type": "Input.ChoiceSet",
                                      "choices": [
                                          {
                                              "title": "Yes",
                                              "value": "Yes"
                                          },
                                          {
                                              "title": "No",
                                              "value": "No"
                                          }
                                      ],
                                      "style": "expanded",
                                      "id": "Question1"
                                  }
                              ],
                              "horizontalAlignment": "Center"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "External or Internal Users",
                              "wrap": true,
                              "id": "Text2"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question2"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Data ownership, type or source",
                              "wrap": true,
                              "id": "Text3"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question3"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Data storage, backup, restore, archival or disaster recovery",
                              "wrap": true,
                              "id": "Text4"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question4"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Data access, performance and sharing",
                              "wrap": true,
                              "id": "Text5"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question5"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Application functionality or capability you intend to use",
                              "wrap": true,
                              "id": "Text6"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question6"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Application interfaces or integration methods",
                              "wrap": true,
                              "id": "Text7"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question7"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Application, server, client or hardware platform",
                              "wrap": true,
                              "id": "Text8"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question8"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Database, network or storage architecture",
                              "wrap": true,
                              "id": "Text9"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question9"
                          }
                      ]
                  }
              ]
          },
          {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Development, test, operations management or software configurations",
                              "wrap": true,
                              "id": "Text10"
                          }
                      ]
                  },
                  {
                      "type": "Column",
                      "width": "auto",
                      "items": [
                          {
                              "type": "Input.ChoiceSet",
                              "placeholder": "Placeholder text",
                              "choices": [
                                {
                                    "title": "Yes",
                                    "value": "Yes"
                                },
                                {
                                    "title": "No",
                                    "value": "No"
                                }
                              ],
                              "style": "expanded",
                              "id": "Question10"
                          }
                      ]
                  }
              ]
          }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {
                    "action": "createRAWVersionUpgrade"
                }
            }
        ]
    });
     }

     createRAW3ArchitectureNew() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
                  {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "What Category Does this Request Fall Into?",
                                "wrap": true
                            }
                          ],
                          "width": "stretch"
                      }
                  ]
              },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Select an Option",
                                  "isSubtle": true,
                                  "wrap": true
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Application Approval",
                  "title": "Application Approval",
                  "data":{
                        "action": "createRAW3ArchitectureNew",
                        "option": "Application Approval"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Custom Solution Approval",
                  "title": "Custom Solution Approval",
                  "data":{
                        "action": "createRAW3ArchitectureNew",
                        "option": "Custom Solution Approval"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Policy Approval",
                  "title": "Policy Approval",
                  "data":{
                        "action": "createRAW3ArchitectureNew",
                        "option": "Policy Approval"
                  }
              }
          ]
      }
      );
     }

     createRAW3ArchitectureChange() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
                  {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "What Category Does this Request Fall Into?",
                                "wrap": true
                            }
                          ],
                          "width": "stretch"
                      }
                  ]
              },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Select an Option",
                                  "isSubtle": true,
                                  "wrap": true
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Version Upgrade",
                  "title": "Version Upgrade",
                  "data":{
                        "action": "createRAW3ArchitectureChange",
                        "option": "Version Upgrade"
                  }
              }
          ]
      }
      );
     }

     createRAW4ArchNewSoftApprovalLicenseVendor() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
                  {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "Whats the Vendor Name?",
                                "wrap": true
                            }
                          ],
                          "width": "stretch"
                      }
                  ]
              },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Enter the Vendor Name",
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "vendorName",
                                  "placeholder": "Name of the Vendor"
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Submit",
                  "title": "Submit",
                  "data":{
                        "action": "createRAW4ArchNewSoftApprovalLicenseVendor"
                  }
              }
          ]
      }
      );
     }


     createRAW4ArchNewSoftApprovalLicenseVendorDesc(vendorName,vendorDesc,vendorWiki) {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
              {
               "type": "TextBlock",
               "text": vendorName,
               "weight": "bolder",
               "isSubtle": false
              },
              {
                "type": "TextBlock",
                "text": vendorDesc,
                "weight": "bolder",
                "size": "medium",
                "wrap": true,
                "separator": true
              },
              {
                "type": "TextBlock",
                "text": vendorWiki,
                "wrap": true
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": 'This One',
                  "title": 'This One',
                  "data":{
                        "action": "createRAW4ArchNewSoftApprovalLicenseVendorDesc",
                        "wiki": vendorWiki,
                        "desc": vendorDesc

                  }
              }
          ]
      }
      );
     }



     createRAW4ArchNewSoftApprovalLicenseName() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
                  {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "Whats the Application's Name?",
                                "wrap": true
                            }
                          ],
                          "width": "stretch"
                      }
                  ]
              },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Enter the Application Name",
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "applicationName",
                                  "placeholder": "Name of Application"
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Submit",
                  "title": "Submit",
                  "data":{
                        "action": "createRAW4ArchNewSoftApprovalLicenseName"
                  }
              }
          ]
      }
      );
     }

     createRAW4ArchNewSoftApprovalLicenseNameDesc(appName,appDesc,appWiki) {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
              {
               "type": "TextBlock",
               "text": appName,
               "weight": "bolder",
               "isSubtle": false
              },
              {
                "type": "TextBlock",
                "text": appDesc,
                "weight": "bolder",
                "size": "medium",
                "wrap": true,
                "separator": true
              },
              {
                "type": "TextBlock",
                "text": appWiki,
                "wrap": true
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": 'This One',
                  "title": 'This One',
                  "data":{
                        "action": "createRAW4ArchNewSoftApprovalLicenseNameDesc",
                        "wiki": appWiki,
                        "desc": appDesc

                  }
              }
          ]
      }
      );
     }

     createVendorAppProfile(vendorName,vendorDesc,vendorWebsite, vendorAppName,vendorAppDesc,vendorAppWebsite,vendorAppNumEmployees,vendorAppType,vendorAppTradedAs,vendorAppISIN,vendorAppIndustry,vendorAppProducts,vendorAppServices,vendorAppFounded,vendorAppFounder,vendorAppHQLocation,vendorAppHQLocationCity,vendorAppHQLocationCountry,vendorAppAreaServed,vendorAppKeyPeople,vendorAppAuthor,vendorAppDeveloper,vendorAppFamily,vendorAppWorkingState,vendorAppSourceModel,vendorAppRTMDate,vendorAppGADate,vendorAppReleased,vendorAppLatestVersion,vendorAppLatestReleaseDate,vendorAppProgrammingLanguage,vendorAppOperatingSystem,vendorAppPlatform,vendorAppSize,vendorAppLanguage,vendorAppGenre,vendorAppPreviewVersion,vendorAppPreviewDate,vendorAppMarketingTarget,vendorAppUpdateModel,vendorAppSupportedPlatforms,vendorAppKernelType,vendorAppUI,vendorAppLicense,vendorAppPrecededBy,vendorAppSucceededBy,vendorAppSupportStatus) {
       return CardFactory.adaptiveCard({
          "type": "AdaptiveCard",
          "body": [
              {
                  "type": "TextBlock",
                  "size": "Medium",
                  "weight": "Bolder",
                  "text": "Vendor / Application Profile"
              },
              {
                "type": "FactSet",
                "facts": [
                  {
                    "title": "VENDOR INFO:",
                    "value": "--VENDOR INFO--",
                    "wrap": true,
                    "separator": true
                  },
                  {
                    "title": "Vendor Name:",
                    "value": vendorName,
                    "wrap": true
                  },
                  {
                    "title": "Website:",
                    "value": vendorWebsite,
                    "wrap": true,
                  },
                  {
                    "title": "Vendor Description:",
                    "value": "",
                    "wrap": true,
                  },
                ]
              },
              {
                "type": "TextBlock",
                "text": vendorDesc,
                "wrap": true
              },
              {
                "type": "FactSet",
                "facts": [
                  {
                    "title": "Number of Employees:",
                    "value": vendorAppNumEmployees,
                    "wrap": true,
                  },
                  {
                    "title": "Type:",
                    "value": vendorAppType,
                    "wrap": true,
                  },
                  {
                    "title": "Traded As:",
                    "value": vendorAppTradedAs,
                    "wrap": true,
                  },
                  {
                    "title": "ISIN:",
                    "value": vendorAppISIN,
                    "wrap": true,
                  },
                  {
                    "title": "Industry:",
                    "value": vendorAppIndustry,
                    "wrap": true,
                  },
                  {
                    "title": "Products:",
                    "value": vendorAppProducts,
                    "wrap": true,
                  },
                  {
                    "title": "Services:",
                    "value": vendorAppServices,
                    "wrap": true,
                  },
                  {
                    "title": "Founded:",
                    "value": vendorAppFounded,
                    "wrap": true,
                  },
                  {
                    "title": "Founder:",
                    "value": vendorAppFounder,
                    "wrap": true,
                  },
                  {
                    "title": "HQ Location:",
                    "value": vendorAppHQLocation,
                    "wrap": true,
                  },
                  {
                    "title": "HQ Location City:",
                    "value": vendorAppHQLocationCity,
                    "wrap": true,
                  },
                  {
                    "title": "HQ Location Country:",
                    "value": vendorAppHQLocationCountry,
                    "wrap": true,
                  },
                  {
                    "title": "Area Served:",
                    "value": vendorAppAreaServed,
                    "wrap": true,
                  },
                  {
                    "title": "Key People:",
                    "value": vendorAppKeyPeople,
                    "wrap": true,
                  },
                    ]
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "title": "APPLICATION INFO:",
                        "value": "--APPLICATION INFO--",
                        "wrap": true,
                        "separator": true
                      },
                      {
                        "title": "Application Name:",
                        "value": vendorAppName,
                        "wrap": true
                      },
                      {
                        "title": "Website:",
                        "value": vendorAppWebsite,
                        "wrap": true,
                      },
                      {
                        "title": "Application Description:",
                        "value": "",
                        "wrap": true,
                      },
                    ]
                  },
                  {
                    "type": "TextBlock",
                    "text": vendorAppDesc,
                    "wrap": true
                  },
                  {
                    "type": "FactSet",
                    "facts": [
                      {
                        "title": "Author:",
                        "value": vendorAppAuthor,
                        "wrap": true,
                      },
                      {
                        "title": "Developer:",
                        "value": vendorAppDeveloper,
                        "wrap": true,
                      },
                      {
                        "title": "Application Family:",
                        "value": vendorAppFamily,
                        "wrap": true,
                      },
                      {
                        "title": "Working State:",
                        "value": vendorAppWorkingState,
                        "wrap": true,
                      },
                      {
                        "title": "Source Model:",
                        "value": vendorAppSourceModel,
                        "wrap": true,
                      },
                      {
                        "title": "RTM Date:",
                        "value": vendorAppRTMDate,
                        "wrap": true,
                      },
                      {
                        "title": "GA Date:",
                        "value": vendorAppGADate,
                        "wrap": true,
                      },
                      {
                        "title": "Released:",
                        "value": vendorAppReleased,
                        "wrap": true,
                      },
                      {
                        "title": "Latest Version:",
                        "value": vendorAppLatestVersion,
                        "wrap": true,
                      },
                      {
                        "title": "Latest Release Date:",
                        "value": vendorAppLatestReleaseDate,
                        "wrap": true,
                      },
                      {
                        "title": "Programming Language:",
                        "value": vendorAppProgrammingLanguage,
                        "wrap": true,
                      },
                      {
                        "title": "Operating System:",
                        "value": vendorAppOperatingSystem,
                        "wrap": true,
                      },
                      {
                        "title": "Platform:",
                        "value": vendorAppPlatform,
                        "wrap": true,
                      },
                      {
                        "title": "Size:",
                        "value": vendorAppSize,
                        "wrap": true,
                      },
                      {
                        "title": "Language:",
                        "value": vendorAppLanguage,
                        "wrap": true,
                      },
                      {
                        "title": "Genre:",
                        "value": vendorAppGenre,
                        "wrap": true,
                      },
                      {
                        "title": "Preview Version:",
                        "value": vendorAppPreviewVersion,
                        "wrap": true,
                      },
                      {
                        "title": "Preview Date:",
                        "value": vendorAppPreviewDate,
                        "wrap": true,
                      },
                      {
                        "title": "Marketing Target:",
                        "value": vendorAppMarketingTarget,
                        "wrap": true,
                      },
                      {
                        "title": "Update Model:",
                        "value": vendorAppUpdateModel,
                        "wrap": true,
                      },
                      {
                        "title": "Supported Platforms:",
                        "value": vendorAppSupportedPlatforms,
                        "wrap": true,
                      },
                      {
                        "title": "Kernel Type:",
                        "value": vendorAppKernelType,
                        "wrap": true,
                      },
                      {
                        "title": "UI:",
                        "value": vendorAppUI,
                        "wrap": true,
                      },
                      {
                        "title": "License:",
                        "value": vendorAppLicense,
                        "wrap": true,
                      },
                      {
                        "title": "Preceded By:",
                        "value": vendorAppPrecededBy,
                        "wrap": true,
                      },
                      {
                        "title": "Succeeded By:",
                        "value": vendorAppSucceededBy,
                        "wrap": true,
                      },
                      {
                        "title": "Support Status:",
                        "value": vendorAppSupportStatus,
                        "wrap": true,
                      }
                    ]
                  }
          ],
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "version": "1.0"
      }
      );
     }

     createRAW5ArchNewSoftApproval() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
                  {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "Will this application be used On-Premise, In the Cloud or Both?",
                                "wrap": true
                            }
                          ],
                          "width": "stretch"
                      }
                  ]
              },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Select an Option",
                                  "isSubtle": true,
                                  "wrap": true
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "On Premise Solution",
                  "title": "On Premise Solution",
                  "value": {
                      "option": "On Premise Solution"
                  },
                  "data":{
                        "action": "createRAW5ArchNewSoftApproval",
                        "option": "On Premise Solution"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Cloud Solution",
                  "title": "Cloud Solution",
                  "data":{
                        "action": "createRAW5ArchNewSoftApproval",
                        "option": "Cloud Solution"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Both",
                  "title": "Both, a On Premise Solution and Cloud Solution",
                  "data":{
                        "action": "createRAW5ArchNewSoftApproval",
                        "option": "Both"
                  }
              }
          ]
      }
      );
     }

     createRAW6ArchNewSoftApprovalLicense() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
                  {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "What Best Describes the license type for this Application?",
                                "wrap": true
                            }
                          ],
                          "width": "stretch"
                      }
                  ]
              },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Select an Option",
                                  "isSubtle": true,
                                  "wrap": true
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Free",
                  "title": "Free",
                  "value": {
                      "option": "Free"
                  },
                  "data":{
                        "action": "createRAW6ArchNewSoftApprovalLicense",
                        "option": "Free"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Trial",
                  "title": "Trial",
                  "data":{
                        "action": "createRAW6ArchNewSoftApprovalLicense",
                        "option": "Trial"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Purchase",
                  "title": "Purchase",
                  "data":{
                        "action": "createRAW6ArchNewSoftApprovalLicense",
                        "option": "Purchase"
                  }
              }
          ]
      }
      );
     }




     createRAW7ArchNewSoftApprovalLicenseNameLOB() {
       return CardFactory.adaptiveCard({
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                    {
                    "type": "ColumnSet",
                    "columns": [
                        {
                            "type": "Column",
                            "items": [
                              {
                                  "type": "Image",
                                  "style": "Person",
                                  "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                  "size": "Small"
                              }
                            ],
                            "width": "auto"
                        },
                        {
                            "type": "Column",
                            "items": [
                              {
                                  "type": "TextBlock",
                                  "weight": "Bolder",
                                  "text": "What Line of Business does this request affect?",
                                  "wrap": true
                              }
                            ],
                            "width": "stretch"
                        }
                    ]
                },
                {
                    "type": "TextBlock",
                    "size": "Medium",
                    "weight": "Bolder",
                    "text": "Line of Business"
                },
                {
                    "type": "Input.Toggle",
                    "title": "Pension",
                    "id": "Pension",
                    "wrap": false,
                    "value": "false"
                },
                {
                    "type": "Input.Toggle",
                    "title": "Health",
                    "id": "Health",
                    "wrap": false,
                    "value": "false"
                },
                {
                    "type": "Input.Toggle",
                    "title": "Investment",
                    "id": "Investment",
                    "wrap": false,
                    "value": "false"
                },
                {
                    "type": "Input.Toggle",
                    "title": "Administration",
                    "id": "Administration",
                    "wrap": false,
                    "value": "false"
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit",
                    "data": {
                        "action": "createRAW7ArchNewSoftApprovalLicenseNameLOB"
                    }
                }
            ]
        }
      );
     }

     createRAW4ProjectPhase() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "Image",
                                  "style": "Person",
                                  "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                  "size": "Small"
                              }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "weight": "Bolder",
                                  "text": "Is this Request for Experimental Approval or Production Approval?",
                                  "wrap": true
                              }
                          ],
                          "width": "stretch"
                      }
                  ]
              },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Select an Option",
                                  "isSubtle": true,
                                  "wrap": true
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Experimental Approval",
                  "title": "Experimental Approval",
                  "data":{
                        "action": "createRAW4ProjectPhase",
                        "option": "Experimental Approval"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Production Approval",
                  "title": "Production Approval",
                  "data":{
                        "action": "createRAW4ProjectPhase",
                        "option": "Production Approval"
                  }
              },
          ]
      }
      );
     }

     createFormType() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                        ],
                        "width": "auto"
                    },
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "Tell me more about your request",
                                "wrap": true
                            }
                        ],
                        "width": "stretch"
                    }
                ]
            },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "What is the Title of this Request?",
                                  "isSubtle": true,
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "RequestTitle",
                                  "isMultiline": true,
                                  "placeholder": "Title of Request"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "What is the Business Problem you are trying to solve?",
                                  "isSubtle": true,
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "BusinessProblem",
                                  "isMultiline": true,
                                  "placeholder": "Business Problem"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "What are the high-level requirements you are trying to solve?",
                                  "isSubtle": true,
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "BusinessRequirements",
                                  "isMultiline": true,
                                  "placeholder": "Business Requirements"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "What are the major business benefits that this solution will provide?",
                                  "isSubtle": true,
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "BusinessBenefits",
                                  "isMultiline": true,
                                  "placeholder": "Business Benefits"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "Additional Information?",
                                  "isSubtle": true,
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "AdditionalInfo",
                                  "isMultiline": true,
                                  "placeholder": "Additional Information"
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Submit",
                  "title": "Submit",
                  "data":{
                        "action": "createFormType"
                  }
              }
          ]
      }
      );
     }

     createFormDivisionChiefApproval() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "Image",
                                  "style": "Person",
                                  "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                  "size": "Small"
                              }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "weight": "Bolder",
                                  "text": "Do you have your division chief approval to submit this request?",
                                  "wrap": true
                              }
                          ],
                          "width": "stretch"
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Yes",
                  "title": "Yes",
                  "data":{
                        "action": "createFormDivisionChiefApproval",
                        "option": "Yes"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "No",
                  "title": "No",
                  "data":{
                        "action": "createFormDivisionChiefApproval",
                        "option": "No"
                  }
              }
          ]
      }
      );
     }

     createFormSubmitRAW(rawPurpose, rawType, rawCategory, rawPhase, lineOfBusiness, rawTitle, businessProblem, businessRequirements, businessBenefits, additionalInfo  ) {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                "size": "Small"
                            }
                        ],
                        "width": "auto"
                    },
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "Here's the information I collected about your Request",
                                "wrap": true
                            }
                        ],
                        "width": "stretch",
                        "verticalContentAlignment": "Center"
                    }
                ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Title:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "horizontalAlignment": "Left",
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": rawTitle,
                              "wrap": true
                          }
                      ]
                  }
              ],
              "separator": true
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Request Purpose:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": rawPurpose,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Request Type:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": rawType,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Request Category:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": rawCategory,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Project Phase:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": rawPhase,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Line of Business:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": lineOfBusiness,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Business Problem:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": businessProblem,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Requirements:     ",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": businessRequirements,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Business Benefits:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": businessBenefits,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
            {
              "type": "ColumnSet",
              "columns": [
                  {
                      "type": "Column",
                      "width": "150px",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": "Additional Information:",
                              "weight": "Bolder",
                              "horizontalAlignment": "Right"
                          }
                      ],
                      "verticalContentAlignment": "Center"
                  },
                  {
                      "type": "Column",
                      "width": "stretch",
                      "items": [
                          {
                              "type": "TextBlock",
                              "text": additionalInfo,
                              "wrap": true
                          }
                      ]
                  }
              ]
            },
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "Image",
                                  "style": "Person",
                                  "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                                  "size": "Small"
                              }
                          ],
                          "width": "auto"
                      },
                      {
                          "type": "Column",
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "weight": "Bolder",
                                  "text": "Can I submit this RAW on your behalf?",
                                  "wrap": true
                              }
                          ],
                          "width": "stretch",
                          "verticalContentAlignment": "Center"
                      }
                  ],
                  "separator": true
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Submit",
                  "title": "Submit",
                  "data":{
                        "action": "createFormSubmitRAW",
                        "option": "Submit"
                  }
              },
              {
                  "type": "Action.Submit",
                  "id": "Cancel",
                  "title": "Cancel",
                  "data":{
                        "action": "createFormSubmitRAW",
                        "option": "Cancel"
                  }
              }
          ]
      }
      );
     }

     createForm() {
       return CardFactory.adaptiveCard({
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "type": "AdaptiveCard",
          "version": "1.0",
          "body": [
              {
                  "type": "ColumnSet",
                  "columns": [
                      {
                          "type": "Column",
                          "width": 2,
                          "items": [
                              {
                                  "type": "TextBlock",
                                  "text": "Request for Architecture Work (RAW)",
                                  "weight": "Bolder",
                                  "size": "Medium"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "Submit a New Request",
                                  "isSubtle": true,
                                  "wrap": true
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "{disclaimer}",
                                  "isSubtle": true,
                                  "wrap": true,
                                  "size": "Small"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "Enter Your Name",
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "myName",
                                  "placeholder": "{myName}"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "Enter Your Email",
                                  "wrap": true
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "myEmail",
                                  "placeholder": "{myEmail}",
                                  "style": "Email"
                              },
                              {
                                  "type": "TextBlock",
                                  "text": "Enter Your Telephone"
                              },
                              {
                                  "type": "Input.Text",
                                  "id": "myTel",
                                  "placeholder": "{myTel}",
                                  "style": "Tel"
                              }
                          ]
                      },
                      {
                          "type": "Column",
                          "width": 1,
                          "items": [
                              {
                                  "type": "Image",
                                  "url": "https://gateway.ipfs.io/ipfs/QmXKfQgKVckfbGSMmzHAGAZ3zr1h8yJNrmEuBaJdNsGECs",
                                  "size": "auto"
                              }
                          ]
                      }
                  ]
              }
          ],
          "actions": [
              {
                  "type": "Action.Submit",
                  "id": "Submit",
                  "title": "Submit",
                  "data":{
                        "action": "createRAW"
                  }
              }
          ]
      }
      );
     }


     createLink(text, link) {
       return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "TextBlock",
             "text": text
           }
         ],
         "actions": [
           {
             "type": "Action.OpenUrl",
             "title": "Click Me",
             "url": link
           }
         ]
       });
     }

     createSportCard(dateEvent, homeTeam, homeScore, homeBadge, awayTeam, awayScore, awayBadge) {
       return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
           "type": "AdaptiveCard",
           "version": "1.0",
           "speak": "The",
           "body": [
             {
               "type": "Container",
               "items": [
                 {
                   "type": "ColumnSet",
                   "columns": [
                     {
                       "type": "Column",
                       "width": "auto",
                       "items": [
                         {
                           "type": "Image",
                           "url": awayBadge,
                           "size": "medium"
                         },
                           {
                           "type": "TextBlock",
                           "text": awayTeam,
                           "horizontalAlignment": "center",
                           "weight": "bolder"
                         }
                       ]
                     },
                     {
                       "type": "Column",
                       "width": "stretch",
                       "separator": true,
                       "spacing": "medium",
                       "items": [
                         {
                           "type": "TextBlock",
                           "text": dateEvent,
                           "horizontalAlignment": "center"
                         },
                         {
                           "type": "TextBlock",
                           "text": "Final",
                           "spacing": "none",
                           "horizontalAlignment": "center"
                         },
                         {
                           "type": "TextBlock",
                           "text": awayScore + " - " + homeScore,
                           "size": "extraLarge",
                           "horizontalAlignment": "center"
                         }
                       ]
                     },
                     {
                       "type": "Column",
                       "width": "auto",
                       "separator": true,
                       "spacing": "medium",
                       "items": [
                         {
                           "type": "Image",
                           "url": homeBadge,
                           "size": "medium",
                           "horizontalAlignment": "center"
                         },
                         {
                           "type": "TextBlock",
                           "text": homeTeam,
                           "horizontalAlignment": "center",
                           "weight": "bolder"
                         }
                       ]
                     }
                   ]
                 }
               ]
             }
           ]
       });
     }

     createWeatherCard(cityName, dateTime, cityTemp, cityTempHi, cityTempLo) {
       return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
           "type": "AdaptiveCard",
           "version": "1.0",
           "speak": "The forecast",
           "body": [
             {
               "type": "TextBlock",
               "text": cityName,
               "size": "large",
               "isSubtle": true
             },
             {
               "type": "TextBlock",
               "text": dateTime,
               "spacing": "none"
             },
             {
               "type": "ColumnSet",
               "columns": [
                 {
                   "type": "Column",
                   "width": "auto",
                   "items": [
                     {
                       "type": "Image",
                       "url": "http://messagecardplayground.azurewebsites.net/assets/Mostly%20Cloudy-Square.png",
                       "size": "small"
                     }
                   ]
                 },
                 {
                   "type": "Column",
                   "width": "auto",
                   "items": [
                     {
                       "type": "TextBlock",
                       "text": cityTemp,
                       "size": "extraLarge",
                       "spacing": "none"
                     }
                   ]
                 },
                 {
                   "type": "Column",
                   "width": "stretch",
                   "items": [
                     {
                       "type": "TextBlock",
                       "text": "°F",
                       "weight": "bolder",
                       "spacing": "small"
                     }
                   ]
                 },
                 {
                   "type": "Column",
                   "width": "stretch",
                   "items": [
                     {
                       "type": "TextBlock",
                       "text": "High: " + cityTempHi,
                       "horizontalAlignment": "left"
                     },
                     {
                       "type": "TextBlock",
                       "text": "Low : " + cityTempLo,
                       "horizontalAlignment": "left",
                       "spacing": "none"
                     }
                   ]
                 }
               ]
             }
           ]
       });
     }

     createGifCard() {

       return CardFactory.animationCard(
           '2%',
           [
               { url: 'http://i.imgur.com/ptJ6Ph6.gif' }
           ],
           [],
           {
               subtitle: 'Retirement Formula'
           }
       );
     }

     createImageCard() {

       return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
             {
                 "type": "ColumnSet",
                 "columns": [
                     {
                         "type": "Column",
                         "items": [
                             {
                                 "type": "Image",
                                 "url": "https://cdn.dribbble.com/users/334335/screenshots/3026014/halloween-eblast-header-3.png"
                             }
                         ]
                     }
                 ]
             }
         ]
       });
     }

     createDocumentCard(title, language, keyPhrases, organizations, persons, locations, glossary1, glossary2) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "TextBlock",
             "text": title,
             "weight": "bolder",
             "size": "medium"
           }
         ],
         "actions": [
           {
             "type": "Action.ShowCard",
             "title": "Language",
             "card": {
               "type": "AdaptiveCard",
               "body": [
                 {
                   "type": "TextBlock",
                   "text": "Document Language:",
                   "weight": "bolder",
                   "size": "small",
                   "separator": true
                 },
                 {
                   "type": "TextBlock",
                   "text": language,
                   "size": "small",
                   "wrap": true
                 },
               ]
             }
           },
             {
               "type": "Action.ShowCard",
               "title": "Key Phrases",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Key Phrases:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": keyPhrases + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Organizations",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Organizations",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": organizations + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Persons",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Persons",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": persons + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Locations",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Locations",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": locations + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "CalPERS Glossary",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "CalPERS Glossary",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": glossary1 + "\r",
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": glossary2 + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             }
           ]
       });
     }

     createAppApprovalCard(appName, appDesc, appType, appId, appStatus, appStatusDate, appNote1, appNote2, appNote3) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
            "type": "TextBlock",
            "text": "Application Portfolio",
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": appName,
             "weight": "bolder",
             "size": "medium",
             "separator": true
           },
           {
             "type": "TextBlock",
             "text": appDesc,
             "wrap": true
           }
         ],
         "actions": [
             {
               "type": "Action.ShowCard",
               "title": "Status",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Status",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Status:",
                         "value": appStatus,
                         "wrap": true
                       },
                       {
                         "title": "Status Date:",
                         "value": appStatusDate,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Notes",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Notes",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "1",
                         "value": appNote1,
                         "wrap": false
                       },
                       {
                         "title": "2",
                         "value": appNote2,
                         "wrap": false
                       },
                       {
                         "title": "3",
                         "value": appNote3,
                         "wrap": false
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Additional Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Additional Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "ID:",
                         "value": appId,
                         "wrap": true
                       },
                       {
                         "title": "Type:",
                         "value": appType,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             }
           ]
       });
     }


     createAppInstalledCard(appName, appClass, appId, appInstalled, appCategory, appSubCategory, appStatusDate, appPublisher, appVersion, appEdition, appReleaseDate, appEndOfSales, appEndofLife, appEndOfSupport, appEndofExtendedSupport) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
            "type": "TextBlock",
            "text": "Flexera",
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": appName,
             "weight": "bolder",
             "size": "medium",
             "separator": true
           },
           {
             "type": "TextBlock",
             "text": appClass,
             "wrap": true
           }
         ],
         "actions": [
             {
               "type": "Action.ShowCard",
               "title": "Application Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Application Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Installed:",
                         "value": appInstalled,
                         "wrap": true
                       },
                       {
                         "title": "Category:",
                         "value": appCategory,
                         "wrap": true
                       },
                       {
                         "title": "Sub-Category:",
                         "value": appSubCategory,
                         "wrap": true
                       },
                       {
                         "title": "Flexera ID:",
                         "value": appId,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Vendor Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Vendor Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Publisher",
                         "value": appPublisher,
                         "wrap": false
                       },
                       {
                         "title": "Version",
                         "value": appVersion,
                         "wrap": false
                       },
                       {
                         "title": "Edition",
                         "value": appEdition,
                         "wrap": false
                       },
                       {
                         "title": "Release Date",
                         "value": appReleaseDate,
                         "wrap": false
                       },
                       {
                         "title": "End of Sales",
                         "value": appEndOfSales,
                         "wrap": false
                       },
                       {
                         "title": "End of Life",
                         "value": appEndofLife,
                         "wrap": false
                       },
                       {
                         "title": "End of Support",
                         "value": appEndOfSupport,
                         "wrap": false
                       },
                       {
                         "title": "End of Extended Support",
                         "value": appEndofExtendedSupport,
                         "wrap": false
                       }
                     ]
                   },
                 ]
               }
             }
           ]
       });
     }

     createRAWCard(rawIdTitle, rawName, rawDesc, rawCategory, rawCategoryOther, rawPhase, rawType, rawBizLine, rawSubmitter, rawSubmitterDiv, rawSubmitterUnit, rawOwner, rawOwnerDiv, rawOwnerUnit, rawDateSubmit, rawDateComplete, rawId) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
            "type": "TextBlock",
            "text": rawIdTitle,
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": rawName,
             "weight": "bolder",
             "size": "medium",
             "separator": true
           },
           {
             "type": "TextBlock",
             "text": rawDesc,
             "wrap": true
           }
         ],
         "actions": [
             {
               "type": "Action.ShowCard",
               "title": "Request Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Request Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Request ID",
                         "value": rawId,
                         "wrap": true
                       },
                       {
                         "title": "Date Submitted:",
                         "value": rawDateSubmit,
                         "wrap": true
                       },
                       {
                         "title": "Date Completed:",
                         "value": rawDateComplete,
                         "wrap": true
                       },
                       {
                         "title": "Category",
                         "value": rawCategory,
                         "wrap": true
                       },
                       {
                         "title": "Sub-Category",
                         "value": rawCategoryOther,
                         "wrap": true
                       },
                       {
                         "title": "Phase",
                         "value": rawPhase,
                         "wrap": true
                       },
                       {
                         "title": "Type",
                         "value": rawType,
                         "wrap": true
                       },
                       {
                         "title": "Business Line",
                         "value": rawBizLine,
                         "wrap": true
                       },
                       {
                         "title": "Submitter",
                         "value": rawSubmitter,
                         "wrap": true
                       },
                       {
                         "title": "Submitter Division",
                         "value": rawSubmitterDiv,
                         "wrap": true
                       },
                       {
                         "title": "Submitter Unit",
                         "value": rawSubmitterUnit,
                         "wrap": true
                       },
                       {
                         "title": "Request Owner",
                         "value": rawOwner,
                         "wrap": true
                       },
                       {
                         "title": "Request Division",
                         "value": rawOwnerDiv,
                         "wrap": true
                       },
                       {
                         "title": "Request Unit",
                         "value": rawOwnerUnit,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             }
           ]
       });
     }

     createFinancialCard(financialId, financialTitle, financialDesc, financialYear, financialContact, financialDivision, financialCost, financialApptioCode, financialPriorPO, financialQuantity) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
            "type": "TextBlock",
            "text": "Spending Plan",
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": financialTitle,
             "weight": "bolder",
             "size": "medium",
             "separator": true
           },
           {
             "type": "TextBlock",
             "text": financialDesc,
             "wrap": true
           }
         ],
         "actions": [
             {
               "type": "Action.ShowCard",
               "title": "Purchase Order",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Purchase Order",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Cost:",
                         "value": "$ " + financialCost,
                         "wrap": true
                       },
                       {
                         "title": "Quantity:",
                         "value": financialQuantity,
                         "wrap": true
                       },
                       {
                         "title": "Year:",
                         "value": financialYear,
                         "wrap": true
                       },
                       {
                         "title": "Prior PO#:",
                         "value": financialPriorPO,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Additional Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Additional Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Contact:",
                         "value": financialContact,
                         "wrap": true
                       },
                       {
                         "title": "Division:",
                         "value": financialDivision,
                         "wrap": true
                       },
                       {
                         "title": "Item ID:",
                         "value": financialId,
                         "wrap": true
                       },
                       {
                         "title": "Apptio Code:",
                         "value": financialApptioCode,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             }
           ]
       });
     }

     createGlossaryCard(division, term, description, definedBy, output, related) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
            "type": "TextBlock",
            "text": division + " Business Glossary",
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": term,
             "weight": "bolder",
             "size": "medium",
             "separator": true
           },
           {
             "type": "TextBlock",
             "text": description,
             "wrap": true
           }
         ],
         "actions": [
             {
               "type": "Action.ShowCard",
               "title": "Additional Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Additional Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Defined By:",
                         "value": definedBy,
                         "wrap": true
                       },
                       {
                         "title": "Output:",
                         "value": output,
                         "wrap": true
                       },
                       {
                         "title": "Related Terms:",
                         "value": related,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "View Mind Map",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                          {
                              "type": "TextBlock",
                              "text": "Mind Map",
                              "weight": "Bolder",
                              "horizontalAlignment": "Center",
                              "size": "Large"
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Related",
                                              "wrap": true,
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder"
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Term",
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder"
                                          }
                                      ],
                                      "verticalContentAlignment": "Center",
                                      "horizontalAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Output",
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder"
                                          }
                                      ],
                                      "horizontalAlignment": "Center",
                                      "verticalContentAlignment": "Center"
                                  }
                              ]
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": related,
                                              "horizontalAlignment": "Center",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "Image",
                                              "url": "https://image.flaticon.com/icons/png/512/120/120833.png"
                                          }
                                      ],
                                      "horizontalAlignment": "Center",
                                      "spacing": "None",
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": term,
                                              "horizontalAlignment": "Center",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "Image",
                                              "url": "https://image.flaticon.com/icons/png/512/120/120833.png"
                                          }
                                      ],
                                      "spacing": "None",
                                      "horizontalAlignment": "Center",
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                          "type": "TextBlock",
                                          "text": output,
                                          "horizontalAlignment": "Center",
                                          "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  }
                              ],
                              "separator": true
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "Image",
                                              "url": "https://image.flaticon.com/icons/png/512/141/141993.png",
                                              "horizontalAlignment": "Center",
                                          }
                                      ]
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  }
                              ],
                              "separator": true
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Defined By",
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  }
                              ]
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": definedBy,
                                              "horizontalAlignment": "Center",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center",
                                      "horizontalAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  }
                              ]
                          }
                      ]
               }
             }
           ]
       });
     }

     createReportCard(title, description, owner, designee, approver, division, classification, language, entities, keyPhrases, sentiment) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
            "type": "TextBlock",
            "text": "Cognos Reports",
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": title,
             "weight": "bolder",
             "size": "medium",
             "separator": true
           },
           {
             "type": "TextBlock",
             "text": description,
             "wrap": true
           }
         ],
         "actions": [
             {
               "type": "Action.ShowCard",
               "title": "Additional Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Additional Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Owner:",
                         "value": owner,
                         "wrap": true
                       },
                       {
                         "title": "Designee:",
                         "value": designee,
                         "wrap": true
                       },
                       {
                         "title": "Approver:",
                         "value": approver,
                         "wrap": true
                       },
                       {
                         "title": "Division:",
                         "value": division,
                         "wrap": true
                       },
                       {
                         "title": "Classification:",
                         "value": classification,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Text Analytics",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Text Analytics",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Text Language:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": language,
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Entities:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": entities + "\r",
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Key Phrases:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": keyPhrases + "\r",
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Sentiment Score:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": sentiment,
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.OpenUrl",
               "title": "View Report",
               "url": "http://adaptivecards.io"
             }
           ]
       });
     }

     createBotCard(text1, text2) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "ColumnSet",
             "columns": [
               {
                 "type": "Column",
                 "width": "auto",
                 "items": [
                   {
                     "type": "Image",
                     "url": "https://ipfs.globalupload.io/QmVAx4JDi3NK4TbeKBWQmj8uHqn63hNuAG2rzttK7khMsS",
                     "size": "small",
                     "style": "person"
                   }
                 ]
               },
               {
                 "type": "Column",
                 "width": "stretch",
                 "items": [
                   {
                     "type": "TextBlock",
                     "text": text1,
                     "weight": "bolder",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "spacing": "none",
                     "text": text2,
                     "isSubtle": true,
                     "wrap": true
                   }
                 ]
               }
             ]
           }
         ]
       });
     }

     createUserCard(picture, name, division) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "ColumnSet",
             "columns": [
               {
                 "type": "Column",
                 "width": "auto",
                 "items": [
                   {
                     "type": "Image",
                     "url": picture,
                     "size": "small",
                     "style": "person"
                   }
                 ]
               },
               {
                 "type": "Column",
                 "width": "stretch",
                 "items": [
                   {
                     "type": "TextBlock",
                     "text": name,
                     "weight": "bolder",
                     "wrap": true
                   },
                    {
                      "type": "TextBlock",
                      "spacing": "none",
                      "text": division,
                      "isSubtle": true,
                      "wrap": true
                    }
                 ]
               }
             ]
           }
         ]
       });
     }

     createComboListCard(helperText, choiceList, selectorValue) {

     return CardFactory.adaptiveCard({
       "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
       "type": "AdaptiveCard",
       "version": "1.0",
       "body": [
         {
          "type": "TextBlock",
          "text": helperText,
          "weight": "bolder",
          "isSubtle": false
         },
         {
           "type": "Input.ChoiceSet",
           "id": selectorValue,
           "style": "compact",
           "value": "0",
           "separator": true,
           "choices": choiceList
         }
       ],
       "actions": [
         {
           "type": "Action.Submit",
           "id": "submit",
           "title": "Submit",
           "data":{
                 "action": selectorValue
           }
         }
       ]
     });
     }
}

module.exports.DialogHelper = DialogHelper;
