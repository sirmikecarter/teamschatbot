// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {TurnContext, MessageFactory, TeamsInfo, TeamsActivityHandler, CardFactory, ActionTypes, AttachmentLayoutTypes} = require('botbuilder');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer, QnAMaker } = require('botbuilder-ai');
const axios = require('axios');
var arraySort = require('array-sort');
const querystring = require('querystring');
const TextEncoder = require('util').TextEncoder;
var wtf = require('wtf_wikipedia');


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
          createRAW3Category: '',
          createRAW4Phase: '',
          createRAW3ArchitectureNew: '',
          createRAW3ArchitectureChange: '',
          createRAW5ArchNewSoftApproval: '',
          createRAW6ArchNewSoftApprovalLicense: '',
          createRAW4ArchNewSoftApprovalLicenseVendor: '',
          createRAW4ArchNewSoftApprovalLicenseName: '',
          createRAW7ArchNewSoftApprovalLicenseNameLOB: '',
          createFormRequestTitle: '',
          createFormBusinessProblem: '',
          createFormBusinessRequirements: '',
          createFormBusinessBenefits: '',
          createFormAdditionalInfo: '',
          createFormDivisionChiefApproval: '',
          createFormSubmitRAW: '',
          createRAWProjectPhase:'',
          appVendorArray: [],
          appDescArray: [],
          vendorName: 'N/A',
          vendorDesc: 'N/A',
          vendorDesc1: 'N/A',
          vendorWebsite: 'N/A',
          vendorAppName: 'N/A',
          vendorAppDesc: 'N/A',
          vendorAppWebsite: 'N/A',
          vendorAppNumEmployees: 'N/A',
          vendorAppType: 'N/A',
          vendorAppTradedAs: 'N/A',
          vendorAppISIN: 'N/A',
          vendorAppIndustry: 'N/A',
          vendorAppProducts: 'N/A',
          vendorAppServices: 'N/A',
          vendorAppFounded: 'N/A',
          vendorAppFounder: 'N/A',
          vendorAppHQLocation: 'N/A',
          vendorAppHQLocationCity: 'N/A',
          vendorAppHQLocationCountry: 'N/A',
          vendorAppAreaServed: 'N/A',
          vendorAppKeyPeople: 'N/A',
          vendorAppAuthor: 'N/A',
          vendorAppDeveloper: 'N/A',
          vendorAppFamily: 'N/A',
          vendorAppWorkingState: 'N/A',
          vendorAppSourceModel: 'N/A',
          vendorAppRTMDate: 'N/A',
          vendorAppGADate: 'N/A',
          vendorAppReleased: 'N/A',
          vendorAppLatestVersion: 'N/A',
          vendorAppLatestReleaseDate: 'N/A',
          vendorAppProgrammingLanguage: 'N/A',
          vendorAppOperatingSystem: 'N/A',
          vendorAppPlatform: 'N/A',
          vendorAppSize: 'N/A',
          vendorAppLanguage: 'N/A',
          vendorAppGenre: 'N/A',
          vendorAppPreviewVersion: 'N/A',
          vendorAppPreviewDate: 'N/A',
          vendorAppMarketingTarget: 'N/A',
          vendorAppUpdateModel: 'N/A',
          vendorAppSupportedPlatforms: 'N/A',
          vendorAppKernelType: 'N/A',
          vendorAppUI: 'N/A',
          vendorAppLicense: 'N/A',
          vendorAppPrecededBy: 'N/A',
          vendorAppSucceededBy: 'N/A',
          vendorAppSupportStatus: 'N/A'

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

              //console.log(context.activity)

              switch (context.activity.value.action) {

                case 'createFormDivisionChiefApproval':

                  if(context.activity.value.option === 'Yes')
                  {
                    this.state.createFormDivisionChiefApproval = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createRAW1Purpose()] });

                  }

                  if(context.activity.value.option === 'No')
                  {
                    this.state.createFormDivisionChiefApproval = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('RAW requests need division chief approval, session cancelled','')] });

                  }

                break;

              case 'createRAW2TypeArch':

              this.state.createRAW1Purpose = "Architecture Approval"

                    switch (context.activity.value.option) {

                    case 'New':
                    this.state.createRAW2Type = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createRAW3ArchitectureNew()] });

                    break;

                    case 'Change':
                    this.state.createRAW2Type = context.activity.value.option
                    await context.sendActivity({ attachments: [this.dialogHelper.createRAW3ArchitectureChange()] });

                    break;

                    }

                break;

              case 'createRAW2TypeMarket':

              this.state.createRAW1Purpose = "Market Analysis"
              this.state.createRAW2Type = "New"
              this.state.createRAW3Category = "Research, " + context.activity.value.selectedValues
              this.state.createRAW4Phase = "Concept"

              //console.log(context.activity.value.selectedValues)
              await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });

              break;

              case 'createRAW3ArchitectureNew':

                      switch (context.activity.value.option) {

                      case 'Application Approval':
                      this.state.createRAW3Category = context.activity.value.option
                      this.state.createRAW4Phase = "Project Approved"
                      await context.sendActivity({ attachments: [this.dialogHelper.createRAW4ProjectPhase()] });

                      break;

                      case 'Custom Solution Approval':
                      this.state.createRAW3Category = context.activity.value.option
                      this.state.createRAW4Phase = "Project Approved"
                      await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });


                      break;

                      case 'Policy Approval':
                      this.state.createRAW3Category = context.activity.value.option
                      this.state.createRAW4Phase = "Project Approved"
                      await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });


                      break;

                      }

                  break;

                  case 'createRAW3ArchitectureChange':

                          switch (context.activity.value.option) {

                          case 'Version Upgrade':
                          this.state.createRAW3Category = context.activity.value.option
                          await context.sendActivity({ attachments: [this.dialogHelper.createRAWVersionUpgrade()] });

                          break;

                          }

                      break;

                  case 'createRAWVersionUpgrade':

                  this.state.createRAW4Phase = "Operations"

                  //console.log(context.activity.value.selectedValues)
                  await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });

                  break;

                  case 'createRAW4ProjectPhase':

                  this.state.createRAWProjectPhase = context.activity.value.option
                  this.state.createRAW3Category = this.state.createRAW3Category + ", " + this.state.createRAWProjectPhase
                  await context.sendActivity({ attachments: [this.dialogHelper.createRAW4ArchNewSoftApprovalLicenseVendor()] });

                  break;

                  case 'createRAW4ArchNewSoftApprovalLicenseVendor':

                  this.state.createRAW4ArchNewSoftApprovalLicenseVendor = context.activity.value.vendorName

                  // Wikipedia
                  this.state.appVendorArray = []
                  var self = this;
                  var appVendorArray = self.state.appVendorArray.slice();


                  await axios.get('https://en.wikipedia.org/w/api.php?action=opensearch&search='+this.state.createRAW4ArchNewSoftApprovalLicenseVendor+'&namespace=0&format=json',
                          { params: {
                            'api-version': '2019-05-06'
                            },
                          headers: {
                            'ContentType': 'application/json'
                    }

                  }).then(response => {

                    if (response){

                      var itemCount = response.data[1].length;
                       //console.log(response.data)
                       //console.log(response.data.length)
                      for (var i = 0; i < itemCount; i++)
                      {
                            var vendorName = response.data[1][i]
                            var vendorDesc = response.data[2][i]
                            var vendorWiki = response.data[3][i]




                            appVendorArray.push({'vendorName': vendorName, 'vendorDesc': vendorDesc, 'vendorWiki': vendorWiki})



                      }

                      self.state.appVendorArray = appVendorArray
                      //console.log(self.state.appVendorArray)
                   }

                  }).catch((error)=>{
                         console.log(error);
                  });

                  // If No Description is Found

                  for (var i = 0; i < self.state.appVendorArray.length; i++)
                  {

                    if (self.state.appVendorArray[i].vendorDesc === ''){
                      //console.log('Yea nothing here')
                      var wikiString = self.state.appVendorArray[i].vendorWiki.replace("https://en.wikipedia.org/wiki/", "");

                      wtf.fetch(wikiString).then(doc => {

                        var splitDescString = doc.sections('').text().split(/\r?\n/)
                        self.state.vendorDesc1 = splitDescString[0]

                      })

                      await new Promise((resolve, reject) => setTimeout(resolve, 300));

                      self.state.appVendorArray[i].vendorDesc = self.state.vendorDesc1

                    }

                  }

                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Please select the description best describing this vendor','')] });

                  var attachments = [];

                  self.state.appVendorArray.forEach(function(data){

                  var card = this.dialogHelper.createRAW4ArchNewSoftApprovalLicenseVendorDesc(data.vendorName, data.vendorDesc, data.vendorWiki)

                  attachments.push(card);

                  }, this)

                  await context.sendActivity({ attachments: attachments,
                  attachmentLayout: AttachmentLayoutTypes.Carousel });

                  break;

                  case 'createRAW4ArchNewSoftApprovalLicenseVendorDesc':

                  //console.log(context.activity.value.option)
                  this.state.vendorName = 'N/A'
                  this.state.vendorDesc = 'N/A'
                  this.state.vendorWebsite = 'N/A'
                  this.state.vendorAppNumEmployees = 'N/A'
                  this.state.vendorAppType = 'N/A'
                  this.state.vendorAppTradedAs = 'N/A'
                  this.state.vendorAppISIN = 'N/A'
                  this.state.vendorAppIndustry = 'N/A'
                  this.state.vendorAppProducts = 'N/A'
                  this.state.vendorAppServices = 'N/A'
                  this.state.vendorAppFounded = 'N/A'
                  this.state.vendorAppFounder = 'N/A'
                  this.state.vendorAppHQLocation = 'N/A'
                  this.state.vendorAppHQLocationCity = 'N/A'
                  this.state.vendorAppHQLocationCountry = 'N/A'
                  this.state.vendorAppAreaServed = 'N/A'
                  this.state.vendorAppKeyPeople = 'N/A'

                  var wikiString = context.activity.value.wiki

                  var wikiString2 = wikiString.replace("https://en.wikipedia.org/wiki/", "");
                  this.state.vendorName = wikiString2
                  this.state.vendorDesc = context.activity.value.desc
                  //this.state.vendorWebsite = context.activity.value.wiki

                  wtf.fetch(wikiString2).then(doc => {
                    console.log(doc.infoboxes(0).json());
                    console.log('--VENDOR INFORMATION--');

                    console.log('Vendor Name: ' + this.state.vendorName);

                    console.log('Vendor Description: ' + this.state.vendorDesc);

                    if(doc.infoboxes(0).json().website){
                      console.log('Website: ' + doc.infoboxes(0).json().website.text);
                      this.state.vendorWebsite = doc.infoboxes(0).json().website.text
                    }else if(doc.infoboxes(0).json().homepage){
                      console.log('Website: ' + doc.infoboxes(0).json().homepage.text);
                      this.state.vendorWebsite = doc.infoboxes(0).json().homepage.text
                    }else{
                      this.state.vendorWebsite = wikiString
                    }

                    if(doc.infoboxes(0).json().num_employees){
                      console.log('Number of Employees: ' + doc.infoboxes(0).json().num_employees.text);
                      this.state.vendorAppNumEmployees = doc.infoboxes(0).json().num_employees.text
                    }
                    if(doc.infoboxes(0).json().type){
                      console.log('Type: ' + doc.infoboxes(0).json().type.text);
                      this.state.vendorAppType = doc.infoboxes(0).json().type.text
                    }
                    if(doc.infoboxes(0).json().traded_as){
                      console.log('Traded As: ' + doc.infoboxes(0).json().traded_as.text);
                      this.state.vendorAppTradedAs= doc.infoboxes(0).json().traded_as.text
                    }
                    if(doc.infoboxes(0).json().isin){
                      console.log('ISIN: ' + doc.infoboxes(0).json().isin.text);
                      this.state.vendorAppISIN = doc.infoboxes(0).json().isin.text
                    }
                    if(doc.infoboxes(0).json().industry){
                      console.log('Industry: ' + doc.infoboxes(0).json().industry.text);
                      this.state.vendorAppIndustry = doc.infoboxes(0).json().industry.text
                    }
                    if(doc.infoboxes(0).json().products){
                      console.log('Products: ' + doc.infoboxes(0).json().products.text);
                      this.state.vendorAppProducts = doc.infoboxes(0).json().products.text
                    }
                    if(doc.infoboxes(0).json().services){
                      console.log('Services: ' + doc.infoboxes(0).json().services.text);
                      this.state.vendorAppServices = doc.infoboxes(0).json().services.text
                    }
                    if(doc.infoboxes(0).json().founded){
                      console.log('Founded: ' + doc.infoboxes(0).json().founded.text);
                      this.state.vendorAppFounded = doc.infoboxes(0).json().founded.text
                    }
                    if(doc.infoboxes(0).json().founder){
                      console.log('Founder: ' + doc.infoboxes(0).json().founder.text);
                      this.state.vendorAppFounder = doc.infoboxes(0).json().founder.text
                    }
                    if(doc.infoboxes(0).json().hq_location){
                      console.log('HQ Location: ' + doc.infoboxes(0).json().hq_location.text);
                      this.state.vendorAppHQLocation = doc.infoboxes(0).json().hq_location.text
                    }
                    if(doc.infoboxes(0).json().hq_location_city){
                      console.log('HQ Location City: ' + doc.infoboxes(0).json().hq_location_city.text);
                      this.state.vendorAppHQLocationCity = doc.infoboxes(0).json().hq_location_city.text
                    }
                    if(doc.infoboxes(0).json().hq_location_country){
                      console.log('HQ Location Country: ' + doc.infoboxes(0).json().hq_location_country.text);
                      this.state.vendorAppHQLocationCountry = doc.infoboxes(0).json().hq_location_country.text
                    }
                    if(doc.infoboxes(0).json().area_served){
                      console.log('Area Served: ' + doc.infoboxes(0).json().area_served.text);
                      this.state.vendorAppAreaServed = doc.infoboxes(0).json().area_served.text
                    }
                    if(doc.infoboxes(0).json().key_people){
                      console.log('Key People: ' + doc.infoboxes(0).json().key_people.text);
                      this.state.vendorAppKeyPeople = doc.infoboxes(0).json().key_people.text
                    }

                  });

                    await context.sendActivity({ attachments: [this.dialogHelper.createRAW4ArchNewSoftApprovalLicenseName()] });

                  break;

                  case 'createRAW4ArchNewSoftApprovalLicenseName':

                  this.state.createRAW4ArchNewSoftApprovalLicenseName = context.activity.value.applicationName

                          await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexApproved + '/docs?',
                                  { params: {
                                    'api-version': '2019-05-06',
                                    'search': this.state.createRAW4ArchNewSoftApprovalLicenseName
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

                            await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This product has already been approved, session cancelled','')] });

                          }else{

                            // Wikipedia
                            this.state.appDescArray = []
                            var self = this;
                            var appDescArray = self.state.appDescArray.slice();


                            await axios.get('https://en.wikipedia.org/w/api.php?action=opensearch&search='+this.state.createRAW4ArchNewSoftApprovalLicenseName+'&namespace=0&format=json',
                                    { params: {
                                      'api-version': '2019-05-06'
                                      },
                                    headers: {
                                      'ContentType': 'application/json'
                              }

                            }).then(response => {

                              if (response){

                                var itemCount = response.data[1].length;

                                 //console.log(response.data)
                                 // console.log(response.data.length)

                                for (var i = 0; i < itemCount; i++)
                                {

                                      const appName = response.data[1][i]
                                      const appDesc = response.data[2][i]
                                      const appWiki = response.data[3][i]

                                      appDescArray.push({'appName': appName, 'appDesc': appDesc, 'appWiki': appWiki})
                                }

                                self.state.appDescArray = appDescArray

                                // console.log(self.state.appDescArray)

                             }

                            }).catch((error)=>{
                                   console.log(error);
                            });

                            // If No Description is Found

                            for (var i = 0; i < self.state.appDescArray.length; i++)
                            {

                              if (self.state.appDescArray[i].appDesc === ''){
                                //console.log('Yea nothing here')
                                var wikiString = self.state.appDescArray[i].appWiki.replace("https://en.wikipedia.org/wiki/", "");

                                wtf.fetch(wikiString).then(doc => {

                                  var splitDescString = doc.sections('').text().split(/\r?\n/)
                                  self.state.vendorDesc1 = splitDescString[0]

                                })

                                await new Promise((resolve, reject) => setTimeout(resolve, 300));

                                self.state.appDescArray[i].appDesc = self.state.vendorDesc1

                              }

                            }

                            await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Please select the description best describing this application','')] });


                            var attachments = [];

                            self.state.appDescArray.forEach(function(data){

                            var card = this.dialogHelper.createRAW4ArchNewSoftApprovalLicenseNameDesc(data.appName, data.appDesc, data.appWiki)

                            attachments.push(card);

                            }, this)

                            await context.sendActivity({ attachments: attachments,
                            attachmentLayout: AttachmentLayoutTypes.Carousel });

                          }

                break;

                case 'createRAW4ArchNewSoftApprovalLicenseNameDesc':

                //console.log(context.activity.value.option)
                this.state.vendorAppName = 'N/A'
                this.state.vendorAppDesc = 'N/A'
                this.state.vendorAppWebsite = 'N/A'
                this.state.vendorAppAuthor = 'N/A'
                this.state.vendorAppDeveloper = 'N/A'
                this.state.vendorAppFamily = 'N/A'
                this.state.vendorAppWorkingState = 'N/A'
                this.state.vendorAppSourceModel = 'N/A'
                this.state.vendorAppRTMDate = 'N/A'
                this.state.vendorAppGADate = 'N/A'
                this.state.vendorAppReleased = 'N/A'
                this.state.vendorAppLatestVersion = 'N/A'
                this.state.vendorAppLatestReleaseDate = 'N/A'
                this.state.vendorAppProgrammingLanguage = 'N/A'
                this.state.vendorAppOperatingSystem = 'N/A'
                this.state.vendorAppPlatform = 'N/A'
                this.state.vendorAppSize = 'N/A'
                this.state.vendorAppLanguage = 'N/A'
                this.state.vendorAppGenre = 'N/A'
                this.state.vendorAppPreviewVersion = 'N/A'
                this.state.vendorAppPreviewDate = 'N/A'
                this.state.vendorAppMarketingTarget = 'N/A'
                this.state.vendorAppUpdateModel = 'N/A'
                this.state.vendorAppSupportedPlatforms = 'N/A'
                this.state.vendorAppKernelType = 'N/A'
                this.state.vendorAppUI = 'N/A'
                this.state.vendorAppLicense = 'N/A'
                this.state.vendorAppPrecededBy = 'N/A'
                this.state.vendorAppSucceededBy = 'N/A'
                this.state.vendorAppSupportStatus = 'N/A'


                var wikiString = context.activity.value.wiki


                var wikiString2 = wikiString.replace("https://en.wikipedia.org/wiki/", "");
                this.state.vendorAppName = wikiString2
                this.state.vendorAppDesc = context.activity.value.desc

                //https://en.wikipedia.org/api/rest_v1/page/summary/Windows_Server_2016
                //https://en.wikipedia.org/w/api.php?action=parse&page=PeaZip
                //https://en.wikipedia.org/w/api.php?action=parse&format=json&page=PeaZip
                //https://en.wikipedia.org/w/api.php?action=query&prop=revisions&rvprop=content&format=xmlfm&titles=PeaZip&rvsection=0
                //https://en.wikipedia.org/wiki/PeaZip?action=raw
                //https://www.npmjs.com/package/wtf_wikipedia

                var runWiki = await wtf.fetch(wikiString2).then(doc => {

                  console.log(doc.infoboxes(0).json());

                  console.log('--GENERAL INFORMATION--');

                  console.log('Name: ' + this.state.vendorAppName);
                  console.log('Description: ' + this.state.vendorAppDesc);

                  if(doc.infoboxes(0).json().website){
                    console.log('Website: ' + doc.infoboxes(0).json().website.text);
                    this.state.vendorAppWebsite = doc.infoboxes(0).json().website.text
                  }else{
                    this.state.vendorAppWebsite = wikiString
                  }

                  console.log('--VENDOR INFORMATION--');

                  if(doc.infoboxes(0).json().num_employees){
                    console.log('Number of Employees: ' + doc.infoboxes(0).json().num_employees.text);
                    this.state.vendorAppNumEmployees = doc.infoboxes(0).json().num_employees.text
                  }
                  if(doc.infoboxes(0).json().type){
                    console.log('Type: ' + doc.infoboxes(0).json().type.text);
                    this.state.vendorAppType = doc.infoboxes(0).json().type.text
                  }
                  if(doc.infoboxes(0).json().traded_as){
                    console.log('Traded As: ' + doc.infoboxes(0).json().traded_as.text);
                    this.state.vendorAppTradedAs= doc.infoboxes(0).json().traded_as.text
                  }
                  if(doc.infoboxes(0).json().isin){
                    console.log('ISIN: ' + doc.infoboxes(0).json().isin.text);
                    this.state.vendorAppISIN = doc.infoboxes(0).json().isin.text
                  }
                  if(doc.infoboxes(0).json().industry){
                    console.log('Industry: ' + doc.infoboxes(0).json().industry.text);
                    this.state.vendorAppIndustry = doc.infoboxes(0).json().industry.text
                  }
                  if(doc.infoboxes(0).json().products){
                    console.log('Products: ' + doc.infoboxes(0).json().products.text);
                    this.state.vendorAppProducts = doc.infoboxes(0).json().products.text
                  }
                  if(doc.infoboxes(0).json().services){
                    console.log('Services: ' + doc.infoboxes(0).json().services.text);
                    this.state.vendorAppServices = doc.infoboxes(0).json().services.text
                  }
                  if(doc.infoboxes(0).json().founded){
                    console.log('Founded: ' + doc.infoboxes(0).json().founded.text);
                    this.state.vendorAppFounded = doc.infoboxes(0).json().founded.text
                  }
                  if(doc.infoboxes(0).json().founder){
                    console.log('Founder: ' + doc.infoboxes(0).json().founder.text);
                    this.state.vendorAppFounder = doc.infoboxes(0).json().founder.text
                  }
                  if(doc.infoboxes(0).json().hq_location){
                    console.log('HQ Location: ' + doc.infoboxes(0).json().hq_location.text);
                    this.state.vendorAppHQLocation = doc.infoboxes(0).json().hq_location.text
                  }
                  if(doc.infoboxes(0).json().hq_location_city){
                    console.log('HQ Location City: ' + doc.infoboxes(0).json().hq_location_city.text);
                    this.state.vendorAppHQLocationCity = doc.infoboxes(0).json().hq_location_city.text
                  }
                  if(doc.infoboxes(0).json().hq_location_country){
                    console.log('HQ Location Country: ' + doc.infoboxes(0).json().hq_location_country.text);
                    this.state.vendorAppHQLocationCountry = doc.infoboxes(0).json().hq_location_country.text
                  }
                  if(doc.infoboxes(0).json().area_served){
                    console.log('Area Served: ' + doc.infoboxes(0).json().area_served.text);
                    this.state.vendorAppAreaServed = doc.infoboxes(0).json().area_served.text
                  }
                  if(doc.infoboxes(0).json().key_people){
                    console.log('Key People: ' + doc.infoboxes(0).json().key_people.text);
                    this.state.vendorAppKeyPeople = doc.infoboxes(0).json().key_people.text
                  }

                  console.log('--PRODUCT INFORMATION--');

                  // if(doc.infoboxes(0).json().logo){
                  //   console.log('Logo: ' + doc.infoboxes(0).json().logo.text);
                  // }
                  // if(doc.infoboxes(0).json().screenshot){
                  //   console.log('Screenshot: ' + doc.infoboxes(0).json().screenshot.text);
                  // }
                  if(doc.infoboxes(0).json().author){
                    console.log('Author: ' + doc.infoboxes(0).json().author.text);
                    this.state.vendorAppAuthor = doc.infoboxes(0).json().author.text
                  }
                  if(doc.infoboxes(0).json().developer){
                    console.log('Developer: ' + doc.infoboxes(0).json().developer.text);
                    this.state.vendorAppDeveloper = doc.infoboxes(0).json().developer.text
                  }
                  if(doc.infoboxes(0).json().family){
                    console.log('Family: ' + doc.infoboxes(0).json().family.text);
                    this.state.vendorAppFamily = doc.infoboxes(0).json().family.text
                  }
                  if(doc.infoboxes(0).json()['working state']){
                    console.log('Working State: ' + doc.infoboxes(0).json()['working state'].text);
                    this.state.vendorAppWorkingState = doc.infoboxes(0).json()['working state'].text
                  }
                  if(doc.infoboxes(0).json()['source model']){
                    console.log('Source Model: ' + doc.infoboxes(0).json()['source model'].text);
                    this.state.vendorAppSourceModel = doc.infoboxes(0).json()['source model'].text
                  }
                  if(doc.infoboxes(0).json()['rtm date']){
                    console.log('RTM Date: ' + doc.infoboxes(0).json()['rtm date'].text);
                    this.state.vendorAppRTMDate = doc.infoboxes(0).json()['rtm date'].text
                  }
                  if(doc.infoboxes(0).json()['ga date']){
                    console.log('GA Date: ' + doc.infoboxes(0).json()['ga date'].text);
                    this.state.vendorAppGADate = doc.infoboxes(0).json()['ga date'].text
                  }
                  if(doc.infoboxes(0).json().released){
                    console.log('Released: ' + doc.infoboxes(0).json().released.text);
                    this.state.vendorAppReleased = doc.infoboxes(0).json().released.text
                  }
                  if(doc.infoboxes(0).json()['latest release version']){
                    console.log('Latest Version: ' + doc.infoboxes(0).json()['latest release version'].text);
                    this.state.vendorAppLatestVersion = doc.infoboxes(0).json()['latest release version'].text
                  }
                  if(doc.infoboxes(0).json()['latest release date']){
                    console.log('Latest Release Date: ' + doc.infoboxes(0).json()['latest release date'].text);
                    this.state.vendorAppLatestReleaseDate = doc.infoboxes(0).json()['latest release date'].text
                  }
                  if(doc.infoboxes(0).json()['programming language']){
                    console.log('Programming Language: ' + doc.infoboxes(0).json()['programming language'].text);
                    this.state.vendorAppProgrammingLanguage = doc.infoboxes(0).json()['programming language'].text
                  }
                  if(doc.infoboxes(0).json()['operating system']){
                    console.log('Operating System: ' + doc.infoboxes(0).json()['operating system'].text);
                    this.state.vendorAppOperatingSystem = doc.infoboxes(0).json()['operating system'].text
                  }
                  if(doc.infoboxes(0).json().platform){
                    console.log('Platform: ' + doc.infoboxes(0).json().platform.text);
                    this.state.vendorAppPlatform = doc.infoboxes(0).json().platform.text
                  }
                  if(doc.infoboxes(0).json().size){
                    console.log('Size: ' + doc.infoboxes(0).json().size.text);
                    this.state.vendorAppSize = doc.infoboxes(0).json().size.text
                  }
                  if(doc.infoboxes(0).json().language){
                    console.log('Language: ' + doc.infoboxes(0).json().language.text);
                    this.state.vendorAppLanguage = doc.infoboxes(0).json().language.text
                  }
                  if(doc.infoboxes(0).json().genre){
                    console.log('Genre: ' + doc.infoboxes(0).json().genre.text);
                    this.state.vendorAppGenre = doc.infoboxes(0).json().genre.text
                  }
                  if(doc.infoboxes(0).json()['preview version']){
                    console.log('Preview Version: ' + doc.infoboxes(0).json()['preview version'].text);
                    this.state.vendorAppPreviewVersion = doc.infoboxes(0).json()['preview version'].text
                  }
                  if(doc.infoboxes(0).json()['preview date']){
                    console.log('Preview Date: ' + doc.infoboxes(0).json()['preview date'].text);
                    this.state.vendorAppPreviewDate = doc.infoboxes(0).json()['preview date'].text
                  }
                  if(doc.infoboxes(0).json()['marketing target']){
                    console.log('Marketing Target: ' + doc.infoboxes(0).json()['marketing target'].text);
                    this.state.vendorAppMarketingTarget = doc.infoboxes(0).json()['marketing target'].text
                  }
                  if(doc.infoboxes(0).json()['update model']){
                    console.log('Update Model: ' + doc.infoboxes(0).json()['update model'].text);
                    this.state.vendorAppUpdateModel = doc.infoboxes(0).json()['update model'].text
                  }
                  if(doc.infoboxes(0).json()['supported platforms']){
                    console.log('Supported Platforms: ' + doc.infoboxes(0).json()['supported platforms'].text);
                    this.state.vendorAppSupportedPlatforms = doc.infoboxes(0).json()['supported platforms'].text
                  }
                  if(doc.infoboxes(0).json()['kernel type']){
                    console.log('Kernel Type: ' + doc.infoboxes(0).json()['kernel type'].text);
                    this.state.vendorAppKernelType = doc.infoboxes(0).json()['kernel type'].text
                  }
                  if(doc.infoboxes(0).json().ui){
                    console.log('UI: ' + doc.infoboxes(0).json().ui.text);
                    this.state.vendorAppUI = doc.infoboxes(0).json().ui.text
                  }
                  if(doc.infoboxes(0).json().license){
                    console.log('License: ' + doc.infoboxes(0).json().license.text);
                    this.state.vendorAppLicense = doc.infoboxes(0).json().license.text
                  }
                  if(doc.infoboxes(0).json()['preceded by']){
                    console.log('Preceded By: ' + doc.infoboxes(0).json()['preceded by'].text);
                    this.state.vendorAppPrecededBy = doc.infoboxes(0).json()['preceded by'].text
                  }
                  if(doc.infoboxes(0).json()['succeeded by']){
                    console.log('Succeeded By: ' + doc.infoboxes(0).json()['succeeded by'].text);
                    this.state.vendorAppSucceededBy = doc.infoboxes(0).json()['succeeded by'].text
                  }

                  if(doc.infoboxes(0).json()['support status']){
                    console.log('Support Status: ' + doc.infoboxes(0).json()['support status'].text);
                    this.state.vendorAppSupportStatus = doc.infoboxes(0).json()['support status'].text
                  }

                });

                // wait
                await new Promise((resolve, reject) => setTimeout(resolve, 1000));

                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('I found all this information about the Vendor and Application','')] });

                await context.sendActivity({ attachments: [this.dialogHelper.createVendorAppProfile(this.state.vendorName,this.state.vendorDesc,this.state.vendorWebsite,this.state.vendorAppName,this.state.vendorAppDesc,this.state.vendorAppWebsite,this.state.vendorAppNumEmployees,this.state.vendorAppType,this.state.vendorAppTradedAs,this.state.vendorAppISIN,this.state.vendorAppIndustry,this.state.vendorAppProducts,this.state.vendorAppServices,this.state.vendorAppFounded,this.state.vendorAppFounder,this.state.vendorAppHQLocation,this.state.vendorAppHQLocationCity,this.state.vendorAppHQLocationCountry,this.state.vendorAppAreaServed,this.state.vendorAppKeyPeople,this.state.vendorAppAuthor,this.state.vendorAppDeveloper,this.state.vendorAppFamily,this.state.vendorAppWorkingState,this.state.vendorAppSourceModel,this.state.vendorAppRTMDate,this.state.vendorAppGADate,this.state.vendorAppReleased,this.state.vendorAppLatestVersion,this.state.vendorAppLatestReleaseDate,this.state.vendorAppProgrammingLanguage,this.state.vendorAppOperatingSystem,this.state.vendorAppPlatform,this.state.vendorAppSize,this.state.vendorAppLanguage,this.state.vendorAppGenre,this.state.vendorAppPreviewVersion,this.state.vendorAppPreviewDate,this.state.vendorAppMarketingTarget,this.state.vendorAppUpdateModel,this.state.vendorAppSupportedPlatforms,this.state.vendorAppKernelType,this.state.vendorAppUI,this.state.vendorAppLicense,this.state.vendorAppPrecededBy,this.state.vendorAppSucceededBy,this.state.vendorAppSupportStatus)] });


                  //await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Will this application be used On-Premise, In the Cloud or Both?','')] });
                  await context.sendActivity({ attachments: [this.dialogHelper.createRAW5ArchNewSoftApproval()] });


                break;

                case 'createRAW5ArchNewSoftApproval':

                      switch (context.activity.value.option) {

                      case 'On Premise Solution':
                          this.state.createRAW5ArchNewSoftApproval = context.activity.value.option
                          await context.sendActivity({ attachments: [this.dialogHelper.createRAW6ArchNewSoftApprovalLicense()] });

                          break;

                      case 'Cloud Solution':
                          this.state.createRAW5ArchNewSoftApproval = context.activity.value.option
                          this.state.createRAW3Category = this.state.createRAW3Category + ", Cloud Subscription"
                          await context.sendActivity({ attachments: [this.dialogHelper.createRAW6ArchNewSoftApprovalLicense()] });
                          break;

                      case 'Both':
                          this.state.createRAW5ArchNewSoftApproval = context.activity.value.option
                          this.state.createRAW3Category = this.state.createRAW3Category + ", Cloud Subscription"
                          await context.sendActivity({ attachments: [this.dialogHelper.createRAW6ArchNewSoftApprovalLicense()] });
                          break;

                      }

                      break;

                  case 'createRAW6ArchNewSoftApprovalLicense':

                        switch (context.activity.value.option) {

                            case 'Free':
                                this.state.createRAW6ArchNewSoftApprovalLicense = context.activity.value.option
                                this.state.createRAW3Category = this.state.createRAW3Category + ", " + this.state.createRAW6ArchNewSoftApprovalLicense + " License Type"
                                await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });
                                break;

                            case 'Trial':
                                this.state.createRAW6ArchNewSoftApprovalLicense = context.activity.value.option
                                this.state.createRAW3Category = this.state.createRAW3Category + ", " + this.state.createRAW6ArchNewSoftApprovalLicense + " License Type"
                                await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });
                                break;

                            case 'Purchase':
                                this.state.createRAW6ArchNewSoftApprovalLicense = context.activity.value.option
                                this.state.createRAW3Category = this.state.createRAW3Category + ", " + this.state.createRAW6ArchNewSoftApprovalLicense + " License Type"
                                await context.sendActivity({ attachments: [this.dialogHelper.createRAW7ArchNewSoftApprovalLicenseNameLOB()] });
                                break;

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
              }

              if (context.activity.value.Administration === 'true'){
                this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB = this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB + ', ' + 'Administration'
              }

              await context.sendActivity({ attachments: [this.dialogHelper.createFormType()] });

              break;

              case 'createFormType':
              this.state.createFormRequestTitle = context.activity.value.RequestTitle
              this.state.createFormBusinessProblem = context.activity.value.BusinessProblem
              this.state.createFormBusinessRequirements = context.activity.value.BusinessRequirements
              this.state.createFormBusinessBenefits = context.activity.value.BusinessBenefits
              this.state.createFormAdditionalInfo = context.activity.value.AdditionalInfo

              // await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Ok, Heres your info... a ' + this.state.createRAW2Type + ' ' + this.state.createRAW1Purpose + ' request' + ' that is a ' + this.state.createRAW3ArchitectureNew + ' and the name of the product is ' + this.state.createRAW4ArchNewSoftApprovalLicenseName + ' and is a ' + this.state.createRAW5ArchNewSoftApproval + ' and the license type is a ' + this.state.createRAW6ArchNewSoftApprovalLicense + ' and the application affects ' + this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB + ' and this approval is for ' + this.state.createRAWProjectPhase,'')] });
              // await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Business Problem: ' + this.state.createFormBusinessProblem,'')] });
              // await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Requirements: ' + this.state.createFormBusinessRequirements,'')] });
              // await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Benefits: ' + this.state.createFormBusinessBenefits,'')] });
              // await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Additional Information: ' + this.state.createFormAdditionalInfo,'')] });
              //await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Can I submit this RAW on your behalf?','')] });

              await context.sendActivity({ attachments: [this.dialogHelper.createFormSubmitRAW(this.state.createRAW1Purpose, this.state.createRAW2Type, this.state.createRAW3Category, this.state.createRAW4Phase, this.state.createRAW7ArchNewSoftApprovalLicenseNameLOB, this.state.createFormRequestTitle, this.state.createFormBusinessProblem, this.state.createFormBusinessRequirements, this.state.createFormBusinessBenefits, this.state.createFormAdditionalInfo)] });

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
                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Session Cancelled','')] });
              }

            break;
            }

            }else{

            switch (context.activity.text.trim()) {
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

              case 'Software_Installed':

              const searchInstalledTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

              var self = this;

              self.state.appArray = []

              await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexInstalled + '/docs?',
                      { params: {
                        'api-version': '2019-05-06',
                        'search': searchInstalledTerm
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
                        const appClass = response.data.value[i].metadata_classification
                        const appPublisher = response.data.value[i].metadata_publisher
                        const appVersion = response.data.value[i].metadata_version
                        const appEdition = response.data.value[i].metadata_edition
                        const appCategory = response.data.value[i].metadata_softwarecategory
                        const appSubCategory = response.data.value[i].metadata_softwaresubcategory
                        const appInstalled = response.data.value[i].metadata_installed
                        const appReleaseDate = response.data.value[i].metadata_releasedate
                        const appEndOfSales = response.data.value[i].metadata_endofsales
                        const appEndofLife = response.data.value[i].metadata_endoflife
                        const appEndOfSupport = response.data.value[i].metadata_endofsupport
                        const appEndofExtendedSupport = response.data.value[i].metadata_endofextendedsupport
                        const appId = response.data.value[i].metadata_flexeraid

                        itemArray.push({'appScore': appScore, 'appName': appName, 'appClass': appClass, 'appPublisher': appPublisher, 'appVersion': appVersion, 'appEdition': appEdition, 'appCategory': appCategory, 'appSubCategory': appSubCategory, 'appInstalled': appInstalled, 'appReleaseDate': appReleaseDate, 'appEndOfSales': appEndOfSales, 'appEndofLife': appEndofLife, 'appEndOfSupport': appEndOfSupport, 'appEndofExtendedSupport': appEndofExtendedSupport, 'appId': appId})
                  }

                  self.state.appArray = arraySort(itemArray, 'appScore')


               }

              }).catch((error)=>{
                     console.log(error);
              });

              //console.log(self.state.appArray)


              if (self.state.appArray.length > 0){



                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from Flexera related to ' + searchInstalledTerm,'')] });

                var attachments = [];

                this.state.appArray.forEach(function(data){

                var card = this.dialogHelper.createAppInstalledCard(data.appName, data.appClass, data.appId, data.appInstalled, data.appCategory, data.appSubCategory, data.appStatusDate, data.appPublisher, data.appVersion, data.appEdition, data.appReleaseDate, data.appEndOfSales, data.appEndofLife, data.appEndOfSupport, data.appEndofExtendedSupport)

                attachments.push(card);

                }, this)

                await context.sendActivity({ attachments: attachments,
                attachmentLayout: AttachmentLayoutTypes.Carousel });

              }else{

                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Installed Applications Related to Your Search Were Found','')] });

              }


              break;

              case 'Software_Create_RAW':

              //console.log(dispatchResults.text)
              await context.sendActivity({ attachments: [this.dialogHelper.createFormDivisionChiefApproval()] });

              break;

              case 'Software_Financials':

                const searchFinancialTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

                var self = this;

                self.state.appArray = []


                await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexFinancials + '/docs?',
                        { params: {
                          'api-version': '2019-05-06',
                          'search': searchFinancialTerm
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
                          const financialScore = i
                          const financialId = response.data.value[i].metadata_itemid
                          const financialTitle = response.data.value[i].questions[0]
                          const financialDesc = response.data.value[i].answer
                          const financialYear = response.data.value[i].metadata_year
                          const financialContact = response.data.value[i].metadata_contact
                          const financialDivision = response.data.value[i].metadata_division
                          const financialCost = response.data.value[i].metadata_cost
                          const financialApptioCode = response.data.value[i].metadata_apptiocode
                          const financialPriorPO = response.data.value[i].metadata_priorpo
                          const financialQuantity = response.data.value[i].metadata_quantity

                          itemArray.push({'financialScore': financialScore, 'financialId': financialId, 'financialTitle': financialTitle, 'financialDesc': financialDesc, 'financialYear': financialYear, 'financialContact': financialContact, 'financialDivision': financialDivision, 'financialCost': financialCost, 'financialApptioCode': financialApptioCode, 'financialPriorPO': financialPriorPO, 'financialQuantity': financialQuantity})
                    }

                    self.state.appArray = arraySort(itemArray, 'financialScore')


                 }

                }).catch((error)=>{
                       console.log(error);
                });

                //console.log(self.state.appArray)

                if (self.state.appArray.length > 0){



                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from the Spending Plan related to ' + searchFinancialTerm,'')] });

                  var attachments = [];

                  this.state.appArray.forEach(function(data){

                  var card = this.dialogHelper.createFinancialCard(data.financialId, data.financialTitle, data.financialDesc, data.financialYear, data.financialContact, data.financialDivision, data.financialCost, data.financialApptioCode, data.financialPriorPO, data.financialQuantity)

                  attachments.push(card);

                  }, this)

                  await context.sendActivity({ attachments: attachments,
                  attachmentLayout: AttachmentLayoutTypes.Carousel });

                }else{

                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Items in the Spending Plans Are Related to Your Search','')] });

                }


                break;

              case 'Software_RAW':

              var self = this;

              var searchRAWTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

              self.state.appArray = []

              await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexRAW + '/docs?',
                      { params: {
                        'api-version': '2019-05-06',
                        'search': searchRAWTerm
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
                        const rawScore = i
                        const rawIdTitle = response.data.value[i].questions[0]
                        const rawName = response.data.value[i].questions[1]
                        const rawDesc = response.data.value[i].answer
                        const rawCategory = response.data.value[i].metadata_requestcategory
                        const rawCategoryOther = response.data.value[i].metadata_requestcategoryother
                        const rawPhase = response.data.value[i].metadata_requestphase
                        const rawType = response.data.value[i].metadata_requesttype
                        const rawBizLine = response.data.value[i].metadata_businessline
                        const rawSubmitter = response.data.value[i].metadata_submittername
                        const rawSubmitterDiv = response.data.value[i].metadata_submitterdivision
                        const rawSubmitterUnit = response.data.value[i].metadata_submitterunit
                        const rawOwner = response.data.value[i].metadata_owner
                        const rawOwnerDiv = response.data.value[i].metadata_ownerdivision
                        const rawOwnerUnit = response.data.value[i].metadata_ownerunit
                        const rawDateSubmit = response.data.value[i].metadata_datesubmitted
                        const rawDateComplete = response.data.value[i].metadata_datecompleted
                        const rawId = response.data.value[i].metadata_rawid

                        itemArray.push({'rawScore': rawScore, 'rawIdTitle': rawIdTitle, 'rawName': rawName, 'rawDesc': rawDesc, 'rawCategory': rawCategory, 'rawCategoryOther': rawCategoryOther, 'rawPhase': rawPhase, 'rawType': rawType, 'rawBizLine': rawBizLine, 'rawSubmitter': rawSubmitter, 'rawSubmitterDiv': rawSubmitterDiv, 'rawSubmitterUnit': rawSubmitterUnit, 'rawOwner': rawOwner, 'rawOwnerDiv': rawOwnerDiv, 'rawOwnerUnit': rawOwnerUnit, 'rawDateSubmit': rawDateSubmit, 'rawDateComplete': rawDateComplete, 'rawId': rawId})
                  }

                  self.state.appArray = arraySort(itemArray, 'rawScore')


               }

              }).catch((error)=>{
                     console.log(error);
              });

              //console.log(self.state.appArray)

              if (self.state.appArray.length > 0){



                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from the RAW system related to ' + searchRAWTerm,'')] });

                var attachments = [];

                this.state.appArray.forEach(function(data){

                var card = this.dialogHelper.createRAWCard(data.rawIdTitle, data.rawName, data.rawDesc, data.rawCategory, data.rawCategoryOther, data.rawPhase, data.rawType, data.rawBizLine, data.rawSubmitter, data.rawSubmitterDiv, data.rawSubmitterUnit, data.rawOwner, data.rawOwnerDiv, data.rawOwnerUnit, data.rawDateSubmit, data.rawDateComplete, data.rawId)

                attachments.push(card);

                }, this)

                await context.sendActivity({ attachments: attachments,
                attachmentLayout: AttachmentLayoutTypes.Carousel });

              }else{

                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No RAWs Related to Your Search Were Found','')] });

              }


              break;

              case 'Software_Approved':

              var self = this;

              const searchAppApprovedTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

              self.state.appArray = []
              self.state.appNotes = []
              self.state.appArrayFinal = []
              self.state.appStatus = []

              await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexApproved + '/docs?',
                      { params: {
                        'api-version': '2019-05-06',
                        'search': searchAppApprovedTerm
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

                var approveCheck = answerExp2.toLowerCase().includes(String(searchAppApprovedTerm));

                //console.log(approveCheck)

                if (approveCheck === false && self.state.appArrayFinal[1]){
                  answerExp1 = self.state.appArrayFinal[1].appName.toLowerCase().replace("[", "");
                  answerExp2 = answerExp1.toLowerCase().replace("]", "");
                  approveCheck = answerExp2.toLowerCase().includes(String(searchAppApprovedTerm));
                }

                //console.log(approveCheck)

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Current')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchAppApprovedTerm + ' is Approved to Use ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Restricted')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchAppApprovedTerm + ' is Approved to Use but is Restricted. Check the Notes tab for the Restriction Note ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Experimental')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchAppApprovedTerm + ' is Approved to Use but for Experimental Purposes Only ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Retired')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No, it appears ' + searchAppApprovedTerm + ' is Retired and No longer approved to Use ','')] });
                }

                if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Sunset')
                {
                  await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchAppApprovedTerm + ' is Approved but will soon reach end of life ','')] });
                }


                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from our Application Portfolio related to ' + searchAppApprovedTerm,'')] });

                var attachments = [];

                this.state.appArrayFinal.forEach(function(data){

                var card = this.dialogHelper.createAppApprovalCard(data.appName, data.appDesc, data.appType, data.appId, data.appStatus, data.appStatusDate, data.appNote1, data.appNote2, data.appNote3)

                attachments.push(card);

                }, this)

                await context.sendActivity({ attachments: attachments,
                attachmentLayout: AttachmentLayoutTypes.Carousel });

              }else{

                await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Applications Related to Your Search Were Found','')] });

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
                break;
            }
            await next();
          }
        });

    }

}

module.exports.TeamsConversationBot = TeamsConversationBot;
