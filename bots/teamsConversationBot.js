// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {TurnContext, MessageFactory, TeamsInfo, TeamsActivityHandler, CardFactory, ActionTypes, AttachmentLayoutTypes} = require('botbuilder');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer, QnAMaker } = require('botbuilder-ai');
const axios = require('axios');
var arraySort = require('array-sort');
const querystring = require('querystring');
const TextEncoder = require('util').TextEncoder;

const { DialogHelper } = require('./helpers/dialogHelper');

const QNA_TOP_N = 1;
const QNA_CONFIDENCE_THRESHOLD = 0.5;

class TeamsConversationBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.state = {
          termArray: [],
          cityName: '',
          cityTemp: '',
          cityTempHi: '',
          cityTempLo: '',
          appArray: [],
          appArrayFinal: [],
          appNotes: [],
          appStatus: [],
          itemCount: '',
          createRAW1Purpose:'',
          createRAW2Type: '',
          createRAW3ArchitectureNew: '',
          createRAW4ArchNewSoftApproval: '',
          createRAW5ArchNewSoftApprovalLicense: '',
          createRAW6ArchNewSoftApprovalLicenseName: '',
          createRAW7ArchNewSoftApprovalLicenseNameLOB: '',
          createFormBusinessProblem: '',
          createFormBusinessRequirements: '',
          createFormBusinessBenefits: '',
          createFormAdditionalInfo: '',
          createFormDivisionChiefApproval: '',
          createFormSubmitRAW: ''

        };

        const luisApplication = {
            applicationId: process.env.LuisAppId,
            azureRegion: process.env.LuisAPIHostName,
            // CAUTION: Authoring key is used in this example as it is appropriate for prototyping.
            // When implimenting for deployment/production, assign and use a subscription key instead of an authoring key.
            endpointKey: process.env.LuisAPIKey
        };

        const luisPredictionOptions = {
            spellCheck: true,
            bingSpellCheckSubscriptionKey: process.env.BingSpellCheck

        };

        this.qnaRecognizer = new QnAMaker({
            knowledgeBaseId: process.env.QnAKbId,
            endpointKey: process.env.QnAEndpointKey,
            host: process.env.QnAHostname
        });

        this.luisRecognizer = new LuisRecognizer(luisApplication, luisPredictionOptions);
        this.dialogHelper = new DialogHelper();


        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);

            if (context.activity.value){

              //console.log(context.activity.value.action)



              switch (context.activity.value.action) {

              case 'createRAW1Purpose':

                  switch (context.activity.value.option) {



                    case 'Architecture':
                    this.state.createRAW1Purpose = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, an ' + this.state.createRAW1Purpose + ' request','')] });
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Type of Request Is This?','')] });
                    await context.sendActivity({ attachments: [this.dialogHelper.createRAW2Type()] });

                    break;

                    case 'Market Analysis':
                    this.state.createRAW1Purpose = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW1Purpose + ' request','')] });
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Type of Request Is This?','')] });
                    //await context.sendActivity({ attachments: [this.dialogHelper.createRAW2Type()] });

                    break;

                  }

                  break;

              case 'createRAW2Type':

                    switch (context.activity.value.option) {

                    case 'New':
                    this.state.createRAW2Type = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request','')] });
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Category Does this Request Fall Into?','')] });
                    await context.sendActivity({ attachments: [this.dialogHelper.createRAW3ArchitectureNew()] });


                    break;

                    case 'Change':
                    this.state.createRAW2Type = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' to ' + this.state.createRAW1Purpose + ' request','')] });
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Category Does this Request Fall Into?','')] });
                    //await context.sendActivity({ attachments: [this.dialogHelper.createRAW3ArchitectureNew()] });

                    break;

                    }

                break;

              case 'createRAW3ArchitectureNew':

                      switch (context.activity.value.option) {

                      case 'Software Approval':
                      this.state.createRAW3ArchitectureNew = context.activity.value.option
                      await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew,'')] });
                      await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Best Describes this Software?','')] });
                      await context.sendActivity({ attachments: [this.dialogHelper.createRAW4ArchNewSoftApproval()] });

                      break;

                      case 'Custom Solution':
                      this.state.createRAW3ArchitectureNew = context.activity.value.option
                      await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew,'')] });

                      break;

                      case 'Architecture Work':
                      this.state.createRAW3ArchitectureNew = context.activity.value.option
                      await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is for ' + this.state.createRAW3ArchitectureNew,'')] });

                      break;

                      }

                  break;

                case 'createRAW4ArchNewSoftApproval':

                      switch (context.activity.value.option) {

                      case 'On Premise Solution':
                          this.state.createRAW4ArchNewSoftApproval = context.activity.value.option
                          await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval,'')] });
                          await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Best Describes the license type for this Software?','')] });
                          await context.sendActivity({ attachments: [this.dialogHelper.createRAW5ArchNewSoftApprovalLicense()] });
                          break;

                      case 'Cloud Solution':
                          this.state.createRAW4ArchNewSoftApproval = context.activity.value.option
                          await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval,'')] });
                          await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Best Describes the license type for this Software?','')] });
                          await context.sendActivity({ attachments: [this.dialogHelper.createRAW5ArchNewSoftApprovalLicense()] });
                          break;

                      case 'Both':
                          this.state.createRAW4ArchNewSoftApproval = context.activity.value.option
                          await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is for ' + this.state.createRAW3ArchitectureNew + ' and Both, a On Premise Solution and Cloud Solution','')] });
                          await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Best Describes the license type for this Software?','')] });
                          await context.sendActivity({ attachments: [this.dialogHelper.createRAW5ArchNewSoftApprovalLicense()] });
                          break;

                      }

                      break;

                  case 'createRAW5ArchNewSoftApprovalLicense':

                        switch (context.activity.value.option) {

                            case 'Free':
                                this.state.createRAW5ArchNewSoftApprovalLicense = context.activity.value.option
                                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval + ' and the license type is ' + this.state.createRAW5ArchNewSoftApprovalLicense,'')] });
                                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What the Name of the Software?','')] });
                                await context.sendActivity({ attachments: [this.dialogHelper.createRAW6ArchNewSoftApprovalLicenseName()] });
                                break;

                            case 'Trial':
                                this.state.createRAW5ArchNewSoftApprovalLicense = context.activity.value.option
                                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval + ' and the license type is a ' + this.state.createRAW5ArchNewSoftApprovalLicense,'')] });
                                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What the Name of the Software?','')] });
                                await context.sendActivity({ attachments: [this.dialogHelper.createRAW6ArchNewSoftApprovalLicenseName()] });
                                break;

                            case 'Purchase':
                                this.state.createRAW5ArchNewSoftApprovalLicense = context.activity.value.option
                                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval + ' and the license type is a ' + this.state.createRAW5ArchNewSoftApprovalLicense,'')] });
                                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What the Name of the Software?','')] });
                                await context.sendActivity({ attachments: [this.dialogHelper.createRAW6ArchNewSoftApprovalLicenseName()] });
                                break;

                        }

                        break;

              case 'createRAW6ArchNewSoftApprovalLicenseName':

                      //await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval + ' and the license type is a ' + this.state.createRAW5ArchNewSoftApprovalLicense,'')] });

              this.state.createRAW6ArchNewSoftApprovalLicenseName = context.activity.value.softwareName



                      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexApproved + '/docs?',
                              { params: {
                                'api-version': '2019-05-06',
                                'search': this.state.createRAW6ArchNewSoftApprovalLicenseName
                                },
                              headers: {
                                'api-key': process.env.SearchServiceKey,
                                'ContentType': 'application/json'
                        }

                      }).then(response => {

                        if (response){

                          this.state.itemCount = response.data.value.length

                       }

                      }).catch((error)=>{
                             console.log(error);
                      });


                      if (this.state.itemCount > 0){

                        await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This software has already been approved, session cancelled','')] });

                      }else{

                        await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Under Construction: Query Flexera API to get Software / Vendor data','')] });
                        await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval + ' and the license type is a ' + this.state.createRAW5ArchNewSoftApprovalLicense + ' and the name of the software is ' + this.state.createRAW6ArchNewSoftApprovalLicenseName,'')] });
                        await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What Line of Business does this software affect?','')] });
                        await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });
                      }

                      break;

              case 'createRAW7ArchNewSoftApprovalLicenseNameLOB':

              if (context.activity.value.Pension === 'true'){
                //console.log(context.activity)
                this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB = 'Pension'
              }

              if (context.activity.value.Health === 'true'){
                this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB = this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB + ', ' + 'Health'
              }

              if (context.activity.value.Investment === 'true'){
                this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB = this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB + ', ' + 'Investment'
                console.log(context.activity)
              }

              if (context.activity.value.Administration === 'true'){
                this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB = this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB + ', ' + 'Administration'
              }

              await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval + ' and the license type is a ' + this.state.createRAW5ArchNewSoftApprovalLicense + ' and the name of the software is ' + this.state.createRAW6ArchNewSoftApprovalLicenseName + ' and the software affects ' + this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB,'')] });
              await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Tell me more about the business problem youre trying to solve','')] });
              await context.sendActivity({ attachments: [this.dialogHelper.createFormBusinessProblem()] });

              break;

              case 'createFormBusinessProblem':
              this.state.createFormBusinessProblem = context.activity.value.BusinessProblem
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Business Problem: ' + this.state.createFormBusinessProblem,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Tell me more about your requirements','')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createFormBusinessRequirements()] });
                  break;

              case 'createFormBusinessRequirements':

              this.state.createFormBusinessRequirements = context.activity.value.BusinessRequirements

                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Requirements: ' + this.state.createFormBusinessRequirements,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Tell me more about the business benefits','')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createFormBusinessBenefits()] });
                  break;

              case 'createFormBusinessBenefits':

              this.state.createFormBusinessBenefits = context.activity.value.BusinessBenefits

                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Benefits: ' + this.state.createFormBusinessBenefits,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Any additonal information?','')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createFormAdditionalInfo()] });
                  break;

              case 'createFormAdditionalInfo':
              this.state.createFormAdditionalInfo = context.activity.value.AdditionalInfo
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Additional Information: ' + this.state.createFormAdditionalInfo,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Do you have your division chief approval to submit this request?','')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createFormDivisionChiefApproval()] });
                  break;

              case 'createFormDivisionChiefApproval':

                if(context.activity.value.option === 'Yes')
                {
                  this.state.createFormDivisionChiefApproval = context.activity.value.option
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Heres your info','')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and is a ' + this.state.createRAW4ArchNewSoftApproval + ' and the license type is a ' + this.state.createRAW5ArchNewSoftApprovalLicense + ' and the name of the software is ' + this.state.createRAW6ArchNewSoftApprovalLicenseName + ' and the software affects ' + this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Business Problem: ' + this.state.createFormBusinessProblem,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Requirements: ' + this.state.createFormBusinessRequirements,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Benefits: ' + this.state.createFormBusinessBenefits,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Additional Information: ' + this.state.createFormAdditionalInfo,'')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Can I submit this RAW on your behalf?','')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createFormSubmitRAW()] });

                }

                if(context.activity.value.option === 'No')
                {
                  this.state.createFormDivisionChiefApproval = context.activity.value.option
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('RAW requests need division chief approval, session cancelled','')] });

                }

              break;

            case 'createFormSubmitRAW':

              if(context.activity.value.option === 'Submit')
              {
                this.state.createFormSubmitRAW = context.activity.value.option
                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('RAW submitted','')] });

              }

              if(context.activity.value.option === 'Cancel')
              {
                this.state.createFormSubmitRAW = context.activity.value.option
                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('session cancelled','')] });

              }

            break;


          }


            }else{

            switch (context.activity.text.trim()) {
            case 'MentionMe':
                await this.mentionActivityAsync(context);
                break;
            case 'UpdateCardAction':
                await this.updateCardActivityAsync(context);
                break;
            case 'Delete':
                await this.deleteCardActivityAsync(context);
                break;
            case 'MessageAllMembers':
                await this.messageAllMembersAsync(context);
                break;
            case 'Select a Term':
                await context.sendActivity(`This is under Construction`);
                break;
            case 'See All Terms':
                await context.sendActivity(`This is under Construction`);
                break;
            case 'Glossary Search':
                await context.sendActivity(`This is under Construction`);
                break;
            default:

            const dispatchResults = await this.luisRecognizer.recognize(context);
            const dispatchTopIntent = LuisRecognizer.topIntent(dispatchResults);

            switch (dispatchTopIntent) {
              case 'General':
                  const qnaResult = await this.qnaRecognizer.generateAnswer(dispatchResults.text, QNA_TOP_N, QNA_CONFIDENCE_THRESHOLD);
                  if (!qnaResult || qnaResult.length === 0 || !qnaResult[0].answer){
                    //await context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                  }else{
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                  }
                  break;

              case 'Software_Create_RAW':

              console.log(dispatchResults.text)
              await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Whats the purpose of this request?','')] });
              await context.sendActivity({ attachments: [this.dialogHelper.createRAW1Purpose()] });



                  break;

              case 'Software_Approved':

              var self = this;

              const searchSoftwareApprovedTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

              //console.log('SOFTWARE NAME: ' + stepContext._info.options.software_name)

              var self = this;

              self.state.appArray = []
              self.state.appNotes = []
              self.state.appArrayFinal = []
              self.state.appStatus = []

              await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexApproved + '/docs?',
                      { params: {
                        'api-version': '2019-05-06',
                        'search': searchSoftwareApprovedTerm
                        },
                      headers: {
                        'api-key': process.env.SearchServiceKey,
                        'ContentType': 'application/json'
                }

              }).then(response => {

                if (response){

                  var itemCount

                  if(response.data.value.length === 1){
                    itemCount = 1
                  }

                  if(response.data.value.length === 2){
                    itemCount = 2
                  }

                  if(response.data.value.length === 3){
                    itemCount = 3
                  }

                  if(response.data.value.length > 3){
                    itemCount = 3
                  }

                  var itemArray = self.state.appArray.slice();

                  for (var i = 0; i < itemCount; i++)
                  {
                        const appScore = i
                        const appName = response.data.value[i].questions[0]
                        const appDesc = response.data.value[i].answer
                        const appType = response.data.value[i].metadata_type
                        const appId = response.data.value[i].metadata_provisionid

                        itemArray.push({'appScore': appScore, 'appName': appName, 'appDesc': appDesc, 'appType': appType, 'appId': appId})
                  }

                  self.state.appArray = arraySort(itemArray, 'appScore')


               }

              }).catch((error)=>{
                     console.log(error);
              });

              //console.log(this.state.appArray)


              var itemArrayFinal = self.state.appArrayFinal.slice();

              for (var i = 0; i < self.state.appArray.length; i++)
              {

                self.state.appNotes = []
                self.state.appStatus = []


              await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexApprovedStatus + '/docs?',
                      { params: {
                        'api-version': '2019-05-06',
                        'search': '*',
                        '$filter': 'metadata_provisionid eq ' + '\'' + self.state.appArray[i].appId + '\''
                        },
                      headers: {
                        'api-key': process.env.SearchServiceKey,
                        'ContentType': 'application/json'
                }

              }).then(response => {

                if (response){

                //console.log(response.data.value[0].@search.score)


                  var noteCount = response.data.value.length

                  var noteArray = self.state.appNotes.slice();
                  var statusArray = self.state.appStatus.slice();

                  for (var i2 = 0; i2 < noteCount; i2++)
                  {
                        const appNotes = response.data.value[i2].answer
                        const appStatus = response.data.value[i2].questions[0]
                        const appStatusDate = response.data.value[i2].metadata_statusdate
                        const appStatusValue = response.data.value[i2].metadata_statusvalue

                        if (noteArray.indexOf(appNotes) === -1 && appNotes !== 'undefined')
                        {
                        noteArray.push(appNotes)
                        }



                        if (appStatusValue === '1')
                        {
                          statusArray.push({'appStatus': appStatus, 'appStatusDate': appStatusDate, 'appStatusValue': appStatusValue})
                        }




                  }

                  self.state.appNotes = noteArray
                  self.state.appStatus = statusArray

                  //console.log(statusArray)

                  itemArrayFinal.push({'appScore': self.state.appArray[i].appScore, 'appName': self.state.appArray[i].appName, 'appDesc': self.state.appArray[i].appDesc, 'appType': self.state.appArray[i].appType, 'appId': self.state.appArray[i].appId, 'appStatus': self.state.appStatus[0].appStatus,'appStatusDate': self.state.appStatus[0].appStatusDate, 'appNote1': self.state.appNotes[0], 'appNote2': self.state.appNotes[1], 'appNote3': self.state.appNotes[2]})



               }

              }).catch((error)=>{
                     console.log(error);
              });


            }

            self.state.appArrayFinal = arraySort(itemArrayFinal, 'appScore')



            //console.log(self.state.appArrayFinal)


              if (self.state.appArrayFinal.length > 0){


                var answerExp1 = self.state.appArrayFinal[0].appName.toLowerCase().replace("[", "");
                var answerExp2 = answerExp1.toLowerCase().replace("]", "");

                var approveCheck = answerExp2.toLowerCase().includes(String(searchSoftwareApprovedTerm));

                //console.log(approveCheck)

                if (approveCheck === false && self.state.appArrayFinal[1]){
                  answerExp1 = self.state.appArrayFinal[1].appName.toLowerCase().replace("[", "");
                  answerExp2 = answerExp1.toLowerCase().replace("]", "");
                  approveCheck = answerExp2.toLowerCase().includes(String(searchSoftwareApprovedTerm));
                }

                //console.log(approveCheck)

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Current')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchSoftwareApprovedTerm + ' is Approved to Use ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Restricted')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchSoftwareApprovedTerm + ' is Approved to Use but is Restricted. Check the Notes tab for the Restriction Note ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Experimental')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchSoftwareApprovedTerm + ' is Approved to Use but for Experimental Purposes Only ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Retired')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No, it appears ' + searchSoftwareApprovedTerm + ' is Retired and No longer approved to Use ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Sunset')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchSoftwareApprovedTerm + ' is Approved but will soon reach end of life ','')] });
                }


                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from our Application Portfolio related to ' + searchSoftwareApprovedTerm,'')] });

                var attachments = [];

                this.state.appArrayFinal.forEach(function(data){

                var card = this.dialogHelper.createAppApprovalCard(data.appName, data.appDesc, data.appType, data.appId, data.appStatus, data.appStatusDate, data.appNote1, data.appNote2, data.appNote3)

                attachments.push(card);

                }, this)

                await context.sendActivity({ attachments: attachments,
                attachmentLayout: AttachmentLayoutTypes.Carousel });

              }else{

                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Software Applications Related to Your Search Were Found','')] });

              }

                  break;
              case 'Weather':

              this.state.cityTemp = ''
              this.state.cityTempHi = ''
              this.state.cityTempLo = ''
              this.state.cityName = ''

              const cityName = dispatchResults.entities.Cities[0]

                  var self = this;

                  await axios.get('https://community-open-weather-map.p.rapidapi.com/weather',
                        { params: {
                          'q': String(cityName),
                          'units': 'imperial'
                          },
                        headers: {
                          'X-RapidAPI-Host': process.env.XRapidAPIHost,
                          'X-RapidAPI-Key': process.env.XRapidAPIKey
                    }

                    }).then(response => {

                      if (response){
                        //console.log(response.data)

                        self.state.cityTemp = response.data.main.temp.toFixed(0)
                        self.state.cityTempHi = response.data.main.temp_max.toFixed(0)
                        self.state.cityTempLo = response.data.main.temp_min.toFixed(0)
                        self.state.cityName = response.data.name


                     }

                    }).catch((error)=>{
                           //console.log(error);
                    });

                    //Use of Date.now() function
                    var d = Date(Date.now());
                    // Converting the number of millisecond in date string
                    var dateTime = d.toString()

                    if(self.state.cityName){
                      await context.sendActivity({ attachments: [this.dialogHelper.createWeatherCard(self.state.cityName, dateTime, self.state.cityTemp,self.state.cityTempHi,self.state.cityTempLo)] });
                    }else{
                      await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...City Name Not Found','')] });
                    }

                  break;
              case 'Glossary':
                  const searchTerm = dispatchResults.entities.Term[0];

                  var self = this;

                  self.state.termArray = []
                  //return await stepContext.beginDialog(SEARCH_GLOSSARY_TERM_DIALOG);

                  if(searchTerm !== undefined){
                    //console.log('Term: ' + searchTerm)
                    var termSearch = "'" + String(searchTerm) + "'"

                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Searching Business Glossary for: ' + searchTerm,'')] });

                    await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
                            { params: {
                              'api-version': '2019-05-06',
                              'search': termSearch
                              },
                            headers: {
                              'api-key': process.env.GlossarySearchServiceKey,
                              'ContentType': 'application/json'
                      }

                    }).then(response => {

                      if (response){

                        var itemCount = response.data.value.length

                        if (itemCount >= 10){
                          itemCount = 9
                        }

                        var itemArray = self.state.termArray.slice();

                        for (var i = 0; i < itemCount; i++)
                        {
                              const glossaryTerm = response.data.value[i].questions[0]
                              const glossaryDescription = response.data.value[i].answer
                              const glossaryDefinedBy = response.data.value[i].metadata_definedby.toUpperCase()
                              const glossaryOutput = response.data.value[i].metadata_output.toUpperCase()
                              const glossaryRelated = response.data.value[i].metadata_related

                              if (itemArray.indexOf(glossaryTerm) === -1)
                              {
                                itemArray.push({'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput, 'related': glossaryRelated})
                              }
                        }

                        self.state.termArray = arraySort(itemArray, 'glossaryterm')


                     }

                    }).catch((error)=>{
                           console.log(error);
                    });

                  //  console.log(self.state.termArray)

                    if (this.state.termArray.length > 0){

                      await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.termArray.length + ' Glossary Terms ','Here are the Results')] });

                      var attachments = [];

                      this.state.termArray.forEach(function(data){

                      var card = this.dialogHelper.createBotCard('...I Found ' + this.state.termArray.length + ' Glossary Terms ','Here are the Results')

                      attachments.push(card);

                      }, this)

                      await context.sendActivity({ attachments: attachments,
                      attachmentLayout: AttachmentLayoutTypes.Carousel });



                    }else{

                      await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

                    }

                  }else{

                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

                  }

                  break;

            }

            //await context.sendActivity(dispatchTopIntent);


                // const value = { count: 0 };
                // const card = CardFactory.heroCard(
                //     'What would you like to do?',
                //     null,
                //     [
                //         {
                //             type: ActionTypes.MessageBack,
                //             title: 'Select a Term',
                //             value: value,
                //             text: 'Select a Term'
                //         },
                //         {
                //             type: ActionTypes.MessageBack,
                //             title: 'See All Terms',
                //             value: null,
                //             text: 'See All Terms'
                //         },
                //         {
                //             type: ActionTypes.MessageBack,
                //             title: 'Glossary Search',
                //             value: null,
                //             text: 'Glossary Search'
                //         }]);
                // await context.sendActivity({ attachments: [card] });
                break;
            }
            await next();
          }
        });

        this.onMembersAddedActivity(async (context, next) => {
            context.activity.membersAdded.forEach(async (teamMember) => {
                if (teamMember.id !== context.activity.recipient.id) {
                    await context.sendActivity(`Welcome to the team ${ teamMember.givenName } ${ teamMember.surname }`);
                }
            });
            await next();
        });
    }

    async mentionActivityAsync(context) {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ mention.text }`);
        replyActivity.entities = [mention];
        await context.sendActivity(replyActivity);
    }

    async updateCardActivityAsync(context) {
        const data = context.activity.value;
        data.count += 1;

        const card = CardFactory.heroCard(
            'Welcome Card',
            `Updated count - ${ data.count }`,
            null,
            [
                {
                    type: ActionTypes.MessageBack,
                    title: 'Update Card',
                    value: data,
                    text: 'UpdateCardAction'
                },
                {
                    type: ActionTypes.MessageBack,
                    title: 'Message all members',
                    value: null,
                    text: 'MessageAllMembers'
                },
                {
                    type: ActionTypes.MessageBack,
                    title: 'Delete card',
                    value: null,
                    text: 'Delete'
                }
            ]);

        card.id = context.activity.replyToId;
        await context.updateActivity({ attachments: [card], id: context.activity.replyToId, type: 'message' });
    }

    async deleteCardActivityAsync(context) {
        await context.deleteActivity(context.activity.replyToId);
    }

    async messageAllMembersAsync(context) {
        const members = await TeamsInfo.getMembers(context);

        members.forEach(async (teamMember) => {
            const message = MessageFactory.text(`Hello ${ teamMember.givenName } ${ teamMember.surname }. I'm a Teams conversation bot.`);

            var ref = TurnContext.getConversationReference(context.activity);
            ref.user = teamMember;

            await context.adapter.createConversation(ref,
                async (t1) => {
                    const ref2 = TurnContext.getConversationReference(t1.activity);
                    await t1.adapter.continueConversation(ref2, async (t2) => {
                        await t2.sendActivity(message);
                    });
                });
        });

        await context.sendActivity(MessageFactory.text('All messages have been sent.'));
    }

}

module.exports.TeamsConversationBot = TeamsConversationBot;
