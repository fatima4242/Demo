//Step1 :  get the Opportunity Type from the Response sheet, onto switch case
function getGoogleSheetTemplateFileInfra(opptyType, cloudOpsOppty)
{
 var googleSheetTemplateFile = "";
 var ccRecipients = "";
 var responsibility = "";
  var accountability = "";
  var accountabilitySupport = "";
 var consult = "";
 var informed = "";
 console.log("opptyType",opptyType);
 switch(opptyType)
 {
 case "Landing Zone Essential":

 //todo put a copy of the template sheet as the id here
   googleSheetTemplateFile = DriveApp.getFileById('1iZX4gk0cEp8se2VqfeyAOmnI33z6e2aF_hFKPc7nsiY');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," +"sumedha.shetty@niveussolutions.com" + "," + ""+ "," +"srujana.thakkallapelly@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
    accountabilitySupport = "srujana.thakkallapelly@niveussolutions.com";
    consult = "suhan.shetty@niveussolutions.com";
    informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Landing Zone Standard":
   googleSheetTemplateFile = DriveApp.getFileById('13awzNtn8t1gYt7fmYBthr0rdanAQrQi3YDmuHDWiuy0');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," +"sumedha.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com"+ "," +"srujana.thakkallapelly@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
    accountabilitySupport = "srujana.thakkallapelly@niveussolutions.com";
    consult = "suhan.shetty@niveussolutions.com";
    informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;


  
 case "Landing Zone Enterprise" :
   googleSheetTemplateFile = DriveApp.getFileById('1g9nELDbsCLKMDMMYAc9nyE36YTlguUM9m1WejVL3bw8');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
  case "Brownfield - Landing zone" :
   googleSheetTemplateFile = DriveApp.getFileById('1zMkxirdR1pA1DjNXSJ2V69ceoEgtHwaJcpsurkMEqwU');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," +"prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;

 case "Migration (Brownfield Setup) - Enterprise":
   googleSheetTemplateFile = DriveApp.getFileById('1AT6BsRi1riJK2_scI4No5kq2-9WwhJ3ytUrexAp2wIQ');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" ;
     accountability = "suraj.rao@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;

 case "Migration (Brownfield Setup) - DN":
   googleSheetTemplateFile = DriveApp.getFileById('1ZAJEd9Fupxr-xTN2iCq-t5WVheYI2rUi4YamgLyEPhc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" ;
     accountability = "suraj.rao@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;


 case "Infrastructure Setup (Greenfield) - Enterprise":
   googleSheetTemplateFile = DriveApp.getFileById('1F098qP-s9K478VS1VfjbvYL3IHnyQ7hoa-HpEeOATXY');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
     responsibility = "prathiksha.kamath@niveussolutions.com" ;
     accountability ="meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;

 case "Infrastructure Setup  (Greenfield) - DN":
   googleSheetTemplateFile = DriveApp.getFileById('1F098qP-s9K478VS1VfjbvYL3IHnyQ7hoa-HpEeOATXY');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
     responsibility ="prathiksha.kamath@niveussolutions.com" ;
     accountability ="meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;
 
 case "CICD - Infrastructure Support":
   googleSheetTemplateFile = DriveApp.getFileById('1ZAJEd9Fupxr-xTN2iCq-t5WVheYI2rUi4YamgLyEPhc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
     responsibility = "meghana.mk@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "GCVE/ Managed VMware (Includes Production)":
   googleSheetTemplateFile = DriveApp.getFileById('1I3qs35AjTT-lrT3ip7MHVZNXd4wZvExYOK7a96AQm4I');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" ;
     responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Managed Disaster Recovery Setup (Warm, Inclusive of Infra and DB setup)":
   googleSheetTemplateFile = DriveApp.getFileById('1KE_kF28KPMF1iViddz8ij1coW91ShZKqaJXZ7qFVDzQ');
  //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com"+ "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
   responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Unmanaged Disaster Recovery Setup (Includes Actifio)":
   googleSheetTemplateFile = DriveApp.getFileById('1RpyT5zKfcM49PrN0jj2Z6ZQy8MvrI40IT2Go7esrMmo');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com"+ "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "HashiCorp Terraform Enterprise Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1_-RMgUk4AwhPzXkbiZhG5KD8F2l7Gtq667xoSkdssEg');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com"+ "," +"srujana.thakkallapelly@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
     accountabilitySupport = "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "HashiCorp Vault Enterprise Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1BH6fJ4Kd2_-9dYZMFrttXn5_AO7UPIQidP4RNMRFCjs');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Pixel streaming Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1y0OWFI5RolT_X9u0dqln3b3GAEhxaZAy3oY2Wm_p0yg');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" ;
     responsibility = "meghana.mk@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

 case "Anthos Setup (For Apigee hybrid / On prem private GCP services connectivity)":
   googleSheetTemplateFile = DriveApp.getFileById('1Aqu8YG4uN-Cx9scebrFQwLJuZZ67cwraMJULRdcJrgs');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" ;
      responsibility = "meghana.mk@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Prisma Cloud - Palo Alto setup for AWS/ GCP/ Azure/ Other CSP":
   googleSheetTemplateFile = DriveApp.getFileById('1os5psKV-0hpq9i0VFmsizy9EkXXjV5t0w5YnM2w6SOs');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com";
     responsibility = "";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "File Share Migration":
   googleSheetTemplateFile = DriveApp.getFileById('15PJiHnp8rkTC-Ci0RRMP1ZBceEU6aKqLr-VcD3eBydc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com";
     responsibility = "";
     accountability = "suraj.rao@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "On-Prem Storage Migration to GCS":
   googleSheetTemplateFile = DriveApp.getFileById('1GJvTdr9A5EF1zVjUJnrUalA8ndAUYPaVJDFHe8Yrxjc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com";
    responsibility = "";
     accountability = "suraj.rao@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "NGFW - Palo Alto Creation":
   googleSheetTemplateFile = DriveApp.getFileById('18mTJri_GGLmugoqRjQ1x2MimlFkZV9cGo0OPl6tR1M4');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com";
     responsibility = "";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
   
 case "Citrix based VDI Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1ivxKIhQ9u2IK6VuBYDPK6qpI5N8QAVhlmtl_wVn3MOk');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," + "sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com"+ "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com";
     accountability = "suraj.rao@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

  case "SAP":
   googleSheetTemplateFile = DriveApp.getFileById('1hNwdVGtjW5CjTW11s_lNCalFd1b5clhD7Eui7nC8Az8');
   //ccRecipients="suhan.shetty@niveussolutions.com";
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "";
     accountabilitySupport = "";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

   case "CloudOps":
    if (cloudOpsOppty == 'Cloud Ops - Managed Services Support')
    {
      googleSheetTemplateFile = DriveApp.getFileById('1DmqK4sof3CzouJh1hz1JKYIgv9k9N61z3v8yogtYjzw');
    }
    if (cloudOpsOppty == 'Cloud Ops - Finops Looker Dashboard')
    {
      googleSheetTemplateFile = DriveApp.getFileById('1Ed2KwUU0_ZOozXmjdbQ83AaRGMszPpbldFryVV1i5Oc');
    }
   //ccRecipients="suhan.shetty@niveussolutions.com";
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "";
     accountabilitySupport = "";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

  // case "Custom":
  //  googleSheetTemplateFile = DriveApp.getFileById('1XzDv98DcDso3cF7EOdoiGhU9KAbJ9w53XHEGAzcWH8g');
  // //  ccRecipients="suhan.shetty@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com";
  //    responsibility ="suhan.shetty@niveussolutions.com";
  //    accountability = "";
  //    accountabilitySupport = "";
  //    consult = "suhan.shetty@niveussolutions.com";
  //    informed = "";
  //    ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," + informed;
  //  break;

  default:
  googleSheetTemplateFile = DriveApp.getFileById('1ZAJEd9Fupxr-xTN2iCq-t5WVheYI2rUi4YamgLyEPhc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" ;
     accountability = "suraj.rao@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;
 }
 console.log("getGoogleSheetTemplateFile ccRecipients",ccRecipients);
 return [googleSheetTemplateFile,ccRecipients];
}

//---------------------------------------------------------------------------------
var archDiagramID = '';
var effortSheetID =  '';


//Get TemplateFile for App
function getGoogleSheetTemplateFileApp(opptyType)
{
  var proposal_template = DriveApp.getFileById('175xsTPWIN25R35AXZoVLZJfp-BAOQeKfMw6hxt3J4Z4');
  console.log("opptyType",opptyType);

 if(opptyType == 'Apigee X Setup and Implementation'){
   var template = DriveApp.getFileById('10cB43xeNOCJUdtzXYnEynAGWL5Y6CGhQ0eC7FXLgMXw');
   archDiagramID = '1RM1A-9CHHfGvm_4GbIfjOIftVvhmewW';
   effortSheetID = '1eiVwUFMajQlePL1GWKV6YkqRvPablnWLK4vPIONUs-s';
 }

 else if(opptyType == 'Apigee Hybrid Setup and Implementation'){
   var template = DriveApp.getFileById('1bGBs6i6uPgtyw_4rytZ4q8T70FjAkFwF7mC5Agmaw_U');
   archDiagramID = '1SkmkvnTwc2L5S2WjFKMmzRyW5wu56PVt';
   effortSheetID = '10AlQlVHzr9DiqU2tYXTEVk9umdkPpTRU85vBDO9sM4A';
 }

 else if(opptyType == 'FlutterFlow Web + Mobile Development' || opptyType == 'FlutterFlow WebApp Development' ||      opptyType == 'FlutterFlow MobileApp Development'){
   var template = DriveApp.getFileById('1qe_MbiiVELvZqPh425WSg_hrFimrPWSZJyILQ_0_fHY');
   archDiagramID = '1vTRZdMeadwtWGczHYXym53690AepXUhN';
   effortSheetID = '1C5mjuQGSEH1CiPsIczHhVkdtu9Ozved-yW7F1lo8QoY';
 }

 else{
  var template = DriveApp.getFileById('1qe_MbiiVELvZqPh425WSg_hrFimrPWSZJyILQ_0_fHY');
  archDiagramID = '1K74Xeb92uhKlS6YeSZKzBBMqZhiPP1Y9';
  effortSheetID = '1C5mjuQGSEH1CiPsIczHhVkdtu9Ozved-yW7F1lo8QoY';
 }
  return [template, archDiagramID, effortSheetID,proposal_template];
}
//----------------------------------------------------------------------------------

//Get templateFile for Data
function getGoogleSheetTemplateFileData(opptyType)
{
  console.log("opptyType",opptyType);
  var template = DriveApp.getFileById('1wTYQoGP8Ie8HNqIkGec4lLzybRZQlpvB-nIXwODz93Q');
  return template;
}

//----------------------------------------------------------------------------------

function onFormSubmit(e)
{
 //Get latest response values with their headers

 var rowRange = getResponseSheetHeaderAndValue('Sales to CE');
 console.log("rowRange.rowValues",rowRange.rowValues);
 
//----------------------------------------------------------------------------------

 var vertical = rowRange.rowValues[2];
 console.log("Vertical",vertical);
 var leadingVertical = rowRange.rowValues[3];
 var clientName = rowRange.rowValues[4];
 clientName = convertToTitleCase(clientName);
 console.log("ClientName",clientName);

 //Destination folder
 const destinationFolder = mainDriveFunction(clientName,leadingVertical);
 console.log("destinationFolder",destinationFolder);
 var dest = DriveApp.getFolderById(destinationFolder);
 console.log("destinationFolder",dest);
 // Move additional files uploaded
 moveFiles('1cpbCjAh-T_w6LL_8p4hyCmQ_UvAqd9Jkw3wrkFvULBbhU5F6MasFvm1nzC8n5cDy-RkhQgzq',dest);
 // Move A to A costing file uploaded
 moveFiles('1jRGwibHYlZTpHDP1VwjLipE6LW4L6X5X6C7MqWhrvVld8a6ZyE0792E7myWpFQY6TGbQUQhu',dest);

  // Move the file to the designated folder
  function moveFiles(sourceFileFolder,dest) {
    var sourceFolder = DriveApp.getFolderById(sourceFileFolder); // ID of the source folder
    var files = sourceFolder.getFiles();

    while (files.hasNext()) {
      var file = files.next();
      var fileCopy = file.makeCopy(file.getName(), dest);
      dest.createFile(fileCopy.getBlob());
      DriveApp.getFileById(file.getId()).setTrashed(true); // Use setTrashed to move to trash
      console.log("File", file.getName(), "has been copied and removed from the source folder.");
      // var file = files.next();
      // dest.createFile(file.getBlob());
      // console.log("File", file);
      // sourceFolder.removeFile(file);
    }
  }


 var scopeSheetID = '';
 var infraRegistryID = '';
 scopeSheetID = '19xxhhcp1b97dp9KeLQcx08Gy07dIhfCkSWvRp8HRSyU';

 //Name the Documents
 var scopeSheetName = clientName +' - Scope Sheet Details - First Draft';
 var sowTemplateName = clientName +'- SOW - First Draft'; 
 var proposalTemplateName = clientName +'- Proposal Draft'; 
 var effortSheetName = clientName +'- Effort Sheet'; 
 var archDiagramName = clientName;
 var infraRegistrySheetName = clientName +'- Infrastructure Registry'; 
 var timelineSheetName = clientName +'- Timeline';
 var retroSheetName = clientName + '- Retro Sheet';
 var handBookName = clientName + '- CloudOps Onboarding Handbook';
 var cloudOpsTemplateName = clientName + '-CloudOps (Managed Services)';
 var finOpsTemplateName = clientName + '-Cloud FinOps Dashboard SOW';
 console.log("org sub file name",scopeSheetName);

//Opp Type 
var opptyTypeInfra = rowRange.rowValues[31];
var opptyTypeApp = rowRange.rowValues[21];
var opptyTypeData = rowRange.rowValues[35];
var opptyType = '';
//date of creation
var creationDate = todaysdate(rowRange.rowValues[0]);
console.log("Date", creationDate);
var projectOverview = rowRange.rowValues[11];
var cloudOpsOppty = rowRange.rowValues[47];

console.log("**cloudopsoppty",cloudOpsOppty)
//Choose vertical and oppty type to create deliverables

if(vertical == 'Infra' && opptyTypeInfra == 'CloudOps'){
  //var [googleSheetTemplateFile,ccRecipients] = getGoogleSheetTemplateFileInfra(opptyTypeInfra);{
    var [googleSheetTemplateFile,ccRecipients] = getGoogleSheetTemplateFileInfra(opptyTypeInfra, cloudOpsOppty);
  if(cloudOpsOppty == 'Cloud Ops - Managed Services Support'){
    var ccRecipients = 'infra-ce@niveussolutions.com' 
    createTemplateCloudOps(googleSheetTemplateFile,clientName,dest,opptyType);
    var handBookID = '1lBjoNZW3tim3SKKTqq4A5G0bXIykr6fMiglRze_s5h8'
    createHandBook(handBookID,clientName, dest,handBookName);
  }
  if(cloudOpsOppty == 'Cloud Ops - Finops Looker Dashboard')
  {
    var ccRecipients = 'infra-ce@niveussolutions.com' 
    effortSheetID = '1VTyNuC2MIDduOvBEz7Urej4h63X4fNnbR4D1f07SKXI';
    createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
    createTemplateFinOps(googleSheetTemplateFile,clientName,dest,opptyType);
  }
}
 else if(vertical == 'Infra' || vertical == 'App + Infra' || vertical == 'Infra + Data' || vertical == 'All'){
  opptyType = opptyTypeInfra;
  console.log("sss",opptyType);
  var [googleSheetTemplateFile,ccRecipients] = getGoogleSheetTemplateFileInfra(opptyTypeInfra);
  var ccRecipients = 'infra-ce@niveussolutions.com' 
  if(opptyType == 'Landing Zone Essential' || opptyType == 'Landing Zone Standard' || opptyType == 'Landing Zone Enterprise')
  {
    effortSheetID = '1k19i_WvSkyZitpsq7QWR0xJRso-ereZLRW8RsTmt11o';
  }
  else
  {
      effortSheetID = '1L1qibUUv7EtYz9D8Q4SvY4jZIv5hQbVMWzp-eOmasEo';
  }
  archDiagramID = '1QtKq9-JfR2fSsN7Wpq3pH7M68DgNqgJm';
  infraRegistryID = '1Jkgc00VELqS4HMJx94jWA62N131HPhiHR1WJ-4LOC2c';
  timelineID = '156qKqPgIitJsO_113DWxvky6n4PI9RZgcQPTzRzRW54';
  retroSheetID = '1rlfQdWvuqFKzhucmRzIDzI9mPzPj215W37yiA-Zm8bk'
  createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
  createSOW(googleSheetTemplateFile,clientName,dest,opptyType);
  createInfraRegistry(infraRegistryID,dest,infraRegistrySheetName);
  createTimelineSheet(timelineID,dest,timelineSheetName);
  createRetroSheet(retroSheetID,dest,retroSheetName);
}
 if(vertical == 'Data' || vertical == 'App + Data' || vertical == 'Infra + Data' || vertical == 'All'){
  opptyType = opptyTypeData;
  var googleSheetTemplateFile = getGoogleSheetTemplateFileData(opptyTypeData);
  var ccRecipients = 'ce-data@niveussolutions.com' + "," + 'ce-partnerBU@niveussolutions.com';
  effortSheetID = '1zGxP9-tQlVMo7ejHQyjCt8JQxBow6xJNWjvlqgBNcGg';
  archDiagramID = '1QtKq9-JfR2fSsN7Wpq3pH7M68DgNqgJm';
  //effortSheetID = '1EOizLsYiHidH3Vesm4c9evWvcC8_NNViL16mY5m9ywo'
  createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
  createSOW(googleSheetTemplateFile,clientName,dest,opptyType);
 }
  if(vertical == 'App' || vertical == 'App + Infra' || vertical == 'App + Data' || vertical == 'All'){
  opptyType = opptyTypeApp;
  var [googleSheetTemplateFile,archDiagramID, effortSheetID,googleProposalFile] = getGoogleSheetTemplateFileApp(opptyTypeApp);
  var ccRecipients = 'ce-app-mod@niveussolutions.com';
  createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
  createSOW(googleSheetTemplateFile,clientName,dest,opptyType);
  createProposalDraft(googleProposalFile,clientName,dest,opptyType);
 }

//Create Scope Sheet
 scopeSheetCopy = createScopeSheet(scopeSheetID,dest,scopeSheetName);
  var newscopeSheetID = scopeSheetCopy.getId();
  updateScopeSheetValue(rowRange.rowHeaderValue,rowRange.rowValues,newscopeSheetID);

if(vertical == 'App + Infra'){
  ccRecipients = 'infra-ce@niveussolutions.com' + "," + 'ce-app-mod@niveussolutions.com';
  console.log("ccReceipient 1234 ", ccRecipients);
}
else if(vertical == 'App + Data'){
  ccRecipients = 'ce-data@niveussolutions.com' + "," + 'ce-app-mod@niveussolutions.com';
  console.log("ccReceipient 1234 ", ccRecipients);
}
else if(vertical == 'Infra + Data'){
  ccRecipients = 'ce-data@niveussolutions.com' + "," + 'infra-ce@niveussolutions.com';
}
else if(vertical == 'All'){
  ccRecipients = 'ce-data@niveussolutions.com' + "," + 'infra-ce@niveussolutions.com' + "," +  'ce-app-mod@niveussolutions.com';
}

 
//----------------------------------------------------------------------------------


 //Get SOW
 function createSOW(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(sowTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
 }

 function createProposalDraft(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(proposalTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
 }
 //get Spreadsheet 
function createScopeSheet(scopeSheetID,dest,scopeSheetName){
  const scopeSheet = DriveApp.getFileById(scopeSheetID);
  const scopeSheetCopy = scopeSheet.makeCopy(scopeSheetName, dest);
  return scopeSheetCopy;
  }
//get Effort Sheet
function createEffortSheet(effortSheetID,clientName,dest){
  const effortSheet = DriveApp.getFileById(effortSheetID);
  const effortSheetCopy = effortSheet.makeCopy(effortSheetName, dest);
  const newEffortSheetID = effortSheetCopy.getId();
  updateEffortSheetValues(clientName,newEffortSheetID);
  }
  //Create Infra Registry Sheet
function createInfraRegistry(infraRegistryID,dest,infraRegistrySheetName){
   const infraRegistrySheet = DriveApp.getFileById(infraRegistryID);
   const effortSheetCopy = infraRegistrySheet.makeCopy(infraRegistrySheetName, dest);
}
//Create Timeline Sheet
function createTimelineSheet(timelineID,dest,timelineSheetName){
   const timelineSheet = DriveApp.getFileById(timelineID);
  const timelineSheetCopy = timelineSheet.makeCopy(timelineSheetName, dest);
}

//Create Retro Sheet
function createRetroSheet(retroSheetID,dest,retroSheetName){
  const retroSheet = DriveApp.getFileById(retroSheetID);
  const retroSheetCopy = retroSheet.makeCopy(retroSheetName, dest);
}

// Cloud Ops Template
function createTemplateCloudOps(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(cloudOpsTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
}
//Finops Template
function createTemplateFinOps(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(finOpsTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
}
//Create Handbook
function createHandBook(handBookID,clientName, dest,handBookName){
  console.log("HandbookID",handBookID);
  const handBook = DriveApp.getFileById(handBookID);
  const handBookCopy = handBook.makeCopy(handBookName, dest);
  const newhandBookID = handBookCopy.getId();
  updateHandBookValue(clientName,newhandBookID);
}

 //Arch Diagram
if(opptyTypeInfra != 'CloudOps'){
  console.log("Opp",opptyTypeInfra);
  var archDiagram = DriveApp.getFileById(archDiagramID);
  var archDiagramCopy = archDiagram.makeCopy(archDiagramName, dest);
}
//Update Scope Sheet Value
function updateScopeSheetValue(responseHeaderValues,responseValues,sheetID){
 var ss = SpreadsheetApp.openById(sheetID);
 SpreadsheetApp.setActiveSpreadsheet(ss);
 ss.getRange('B1').setValue(responseValues[0].toString());
 for (var i=1;i<=responseHeaderValues.length;i++){
   //Setup the Response headers
   ss.getRange('A'+ i.toString()).setValue(responseHeaderValues[i-1].toString());
   //Setup the Response ValuescopiedSheetID
   ss.getRange('B'+ i.toString()).setValue(responseValues[i-1].toString());
 }
}
//Update SOW Values
function updateSowValues(clientName,docID,opptyType){
  const doc = DocumentApp.openById(docID);
  const body = doc.getBody();
  body.replaceText('{{Client Name}}',clientName);
  body.replaceText('{{Date of SOW}}',creationDate);
  body.replaceText('{{Project Overview}}',projectOverview);
  body.replaceText('{{Opp Type}}',opptyType);
}

//Update Effort Sheet Values
function updateEffortSheetValues(clientName,sheetID){
  const sheet = SpreadsheetApp.openById(sheetID);
  var textFinder = sheet.createTextFinder("{{Client Name}}");
  textFinder.replaceAllWith(clientName);
}

//Update Handbook Values
function updateHandBookValue(clientName,handBookID){
  const doc = DocumentApp.openById(handBookID);
  const body = doc.getBody();
  body.replaceText('{{Client Name}}',clientName);
  body.replaceText('{{Date of SOW}}',creationDate);
}

 //var recipient_email = rowRange.rowValues[1].toString();
 //sendEmail(org_name, copiedSheetID,recipient_email,mailID,ccRecipients);
 
// const oppType = rowRange.rowValues[31];

 const typeOpp = rowRange.rowValues[8];
 const salesSpocEmail= rowRange.rowValues[1];
 salesSpocName =extractNameFromEmail(salesSpocEmail);
 console.log("SalesSpoc",salesSpocName);
 const hubSpotLink = rowRange.rowValues[6];
 const dealSize = rowRange.rowValues[15];
 var constraints = rowRange.rowValues[20];
 console.log("** Get Date",rowRange.rowValues[18]);
 const nextCall = toDateFmt(rowRange.rowValues[18]);
 console.log("formattedDate = ",nextCall);
 //const problemStatement = rowRange.rowValues[9];
 if(constraints === "")
 constraints = 'Not Applicable';
 console.log("Constraints",constraints);
 var link = 'https://drive.google.com/drive/folders/'+destinationFolder;
 if(vertical == 'All')
 vertical = 'App + Infra + Data';
 ccRecipients = ccRecipients + "," +salesSpocEmail + "," + "salesops@niveussolutions.com" + "," + "suyog.shetty@niveussolutions.com";
 console.log(ccRecipients);
 //ccRecipients = 'fatima.reihab@niveussolutions.com';
   sendEmail(opptyType,vertical,clientName,salesSpocName,typeOpp,projectOverview,hubSpotLink,dealSize,constraints,nextCall,link,ccRecipients);
}

//----------------------------------------------------------------------------------


//Step7 : get the rowValues from the Re+ "," +sponse sheet
function getResponseSheetHeaderAndValue(sheetName){
 const sheet2 = SpreadsheetApp.getActiveSpreadsheet();
 var last_row = sheet2.getLastRow() - 1;
 var range = sheet2.getDataRange();
 var rowHeaderValue = range.getValues()[0];
 var rowValues = range.getValues()[last_row];
 return {rowHeaderValue , rowValues};
}
 
//Step8 : Append the rowValues from the Response sheet to the Scope template sheet


//Sending Email
function sendEmail(typeOfDeal,vertical,client,sales,dealType,projectOverview,hubspotLink,dealSize,constraints,nextCallDate,link,ccRecipients) {
  var data ={typeOfDeal,vertical,client,sales,dealType,projectOverview,hubspotLink,dealSize,constraints,nextCallDate,link};
  var templ = HtmlService.createTemplateFromFile('templateEmail');
  templ.data=data;

  var message = templ.evaluate().getContent();
  console.log("ccReceipient",ccRecipients);
    MailApp.sendEmail({
  //Sending mail based on the verticals
    to: ccRecipients,
    //to: "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com",
    subject: "New enagagement with "+client,
    htmlBody: message,
    cc: 'bharath.yaragatti@niveussolutions.com' 
  });
}

var folder = "NewFolder";

//------------------------------------------------------------------------------------------------------------
//Destination folder
function mainDriveFunction(clientName,leadingVertical) {
  if(leadingVertical == 'App Mod')
    var parentFolderID = '1tX3Q9NuPRZsvEhd2JeArRYNMiXsn_t6A'; 
  else if(leadingVertical == 'Infra Mod')
    var parentFolderID = '1ZIs7pRV8SUb2sO9yyFWXGNwL3UC2cEED';
  else if(leadingVertical == 'Data Mod')
    var parentFolderID = '1tUu9_eFLR56LqF89NEH-W8Mw4XdybOw0';
  else
    console.log("Invalid Vertical");

  const month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  const d = new Date();
  let currMonth = month[d.getMonth()];
  let currYear = d.getFullYear();
  console.log('Year is', currYear);
  var folderID = getFolderByName(parentFolderID,currYear); // Year Folder
  console.log("Folder ID is ", folderID);
  if(folderID==null) {
    createFolder(currYear, parentFolderID); // Created new folder
    folderID = getFolderByName(parentFolderID,currYear);   // Get folder ID of new folder that was created
  }
  parentFolderID = folderID; 
  console.log("Year done");

  console.log('Month is', currMonth);
  folderID = getFolderByName(parentFolderID,currMonth);      // Month Folder
  if(folderID==null) {
    createFolder(currMonth, parentFolderID); // Created new folder
    folderID = getFolderByName(parentFolderID,currMonth);   // Get folder ID of new folder that was created
  }
  parentFolderID = folderID;
  console.log("Month done");

  console.log('clientName is', clientName);
  folderID = getFolderByName(parentFolderID,clientName);      // Client's Name Folder
  if(folderID==null) {
    createFolder(clientName, parentFolderID); // Created new folder
    folderID = getFolderByName(parentFolderID,clientName);   // Get folder ID of new folder that was created
  }
  parentFolderID = folderID;
  console.log("Client done");
  console.log("parentFolderID",parentFolderID);
  Utilities.sleep(5000);
  // if(parentFolderID == null)
  //   {
  //     parentFolderIdNew = parentFolderID;
  //     var i = 3;
  //     while(parentFolderIdNew == null && i>0)
  //     {
  //       Utilities.sleep(3000);
  //       i--;
  //       var parentFolderIdNew1 = checkNullFolder(parentFolderID);
  //       console.log("parentID",parentFolderIdNew1);
  //     }
  //   }
  var folderID = getFolderByName(parentFolderID,clientName); 
  console.log("Client folderID =",folderID);
  return parentFolderID;
  //Copy SOW and other files
  //createDeliverables(folderID);
}

// Create new folder if it doesn't exist
function createFolder(folderName, folderDestination) {
  var parentFolder = DriveApp.getFolderById(folderDestination);
  var folderName = parentFolder.createFolder(folderName);
  console.log("Created Folder - "+folderName+" Folder Destination - "+folderDestination);
}
//mainDriveFunction

// Get folder id by name
function getFolderByName(parentFolderID,folderName) {
  console.log("Parent Folder",parentFolderID);
  console.log("Folder Name", folderName);
  //sometimes error here
  var folders = DriveApp.getFolderById(parentFolderID).getFoldersByName(folderName);
  var folderID = null;
  if (folders.hasNext())
    folderID = folders.next();
  if(folderID != null) {
    var folderIDdestination=folderID.getId();
    console.log("Folder Found - "+folderID+" Folder ID is - "+folderIDdestination);
    return folderID.getId();
  }
  else {
    console.log('Null Folder');
    return null;
  }
}


function toDateFmt(dt_string) {
  console.log("toDateFmt",dt_string);

  var millis = Date.parse(dt_string);
  var date = new Date(millis);

  console.log("extracted date 1",date);
  console.log("extracted date 2",date.getDate());
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  // Return the mainDriveFunctione date in dd-mm-YYYY format
  return `${day}-${month}-${year}`;
}

//Extracting sales spoc name
function extractNameFromEmail(salesSpoc) {
  var name = salesSpoc.match(/^([^@]*)@/)[1]; // Extract the name before the '@' symbol
  name = name.replace('.', ' '); // Replace dots with spaces
  name = name.replace(/\b\w/g, function(l){ return l.toUpperCase() }); // Capitalize the first letter of each word
  return name;
}

//Extracting Date
function todaysdate(timestampValue){
  console.log("Date",timestampValue)
  var date = new Date(timestampValue); // Convert the timestamp to a date object
  var dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-YYYY"); // Format the date object into a string with only the date portion
  return dateString;
}

// Title Case 
function convertToTitleCase(str) {
  const newStr = str.toLowerCase().split(' ')
    .map(w => {
      if (w) {
        return w.charAt(0).toUpperCase() + w.slice(1);
      }
      return '';
    })
    .join(' ');
  console.log(newStr);
  return newStr;
}


//Check if the Parent Folder is Null

function checkNullFolder(parentFolderID)
{
  return parentFolderID;
}//Step1 :  get the Opportunity Type from the Response sheet, onto switch case
function getGoogleSheetTemplateFileInfra(opptyType, cloudOpsOppty)
{
 var googleSheetTemplateFile = "";
 var ccRecipients = "";
 var responsibility = "";
  var accountability = "";
  var accountabilitySupport = "";
 var consult = "";
 var informed = "";
 console.log("opptyType",opptyType);
 switch(opptyType)
 {
 case "Landing Zone Essential":

 //todo put a copy of the template sheet as the id here
   googleSheetTemplateFile = DriveApp.getFileById('1iZX4gk0cEp8se2VqfeyAOmnI33z6e2aF_hFKPc7nsiY');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," +"sumedha.shetty@niveussolutions.com" + "," + ""+ "," +"srujana.thakkallapelly@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
    accountabilitySupport = "srujana.thakkallapelly@niveussolutions.com";
    consult = "suhan.shetty@niveussolutions.com";
    informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Landing Zone Standard":
   googleSheetTemplateFile = DriveApp.getFileById('13awzNtn8t1gYt7fmYBthr0rdanAQrQi3YDmuHDWiuy0');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," +"sumedha.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com"+ "," +"srujana.thakkallapelly@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
    accountabilitySupport = "srujana.thakkallapelly@niveussolutions.com";
    consult = "suhan.shetty@niveussolutions.com";
    informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;


  
 case "Landing Zone Enterprise" :
   googleSheetTemplateFile = DriveApp.getFileById('1g9nELDbsCLKMDMMYAc9nyE36YTlguUM9m1WejVL3bw8');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
  case "Brownfield - Landing zone" :
   googleSheetTemplateFile = DriveApp.getFileById('1zMkxirdR1pA1DjNXSJ2V69ceoEgtHwaJcpsurkMEqwU');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," +"prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;

 case "Migration (Brownfield Setup) - Enterprise":
   googleSheetTemplateFile = DriveApp.getFileById('1AT6BsRi1riJK2_scI4No5kq2-9WwhJ3ytUrexAp2wIQ');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" ;
     accountability = "suraj.rao@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;

 case "Migration (Brownfield Setup) - DN":
   googleSheetTemplateFile = DriveApp.getFileById('1ZAJEd9Fupxr-xTN2iCq-t5WVheYI2rUi4YamgLyEPhc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" ;
     accountability = "suraj.rao@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;


 case "Infrastructure Setup (Greenfield) - Enterprise":
   googleSheetTemplateFile = DriveApp.getFileById('1F098qP-s9K478VS1VfjbvYL3IHnyQ7hoa-HpEeOATXY');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
     responsibility = "prathiksha.kamath@niveussolutions.com" ;
     accountability ="meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;

 case "Infrastructure Setup  (Greenfield) - DN":
   googleSheetTemplateFile = DriveApp.getFileById('1F098qP-s9K478VS1VfjbvYL3IHnyQ7hoa-HpEeOATXY');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
     responsibility ="prathiksha.kamath@niveussolutions.com" ;
     accountability ="meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;
 
 case "CICD - Infrastructure Support":
   googleSheetTemplateFile = DriveApp.getFileById('1ZAJEd9Fupxr-xTN2iCq-t5WVheYI2rUi4YamgLyEPhc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
     responsibility = "meghana.mk@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "GCVE/ Managed VMware (Includes Production)":
   googleSheetTemplateFile = DriveApp.getFileById('1I3qs35AjTT-lrT3ip7MHVZNXd4wZvExYOK7a96AQm4I');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" ;
     responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Managed Disaster Recovery Setup (Warm, Inclusive of Infra and DB setup)":
   googleSheetTemplateFile = DriveApp.getFileById('1KE_kF28KPMF1iViddz8ij1coW91ShZKqaJXZ7qFVDzQ');
  //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com"+ "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
   responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Unmanaged Disaster Recovery Setup (Includes Actifio)":
   googleSheetTemplateFile = DriveApp.getFileById('1RpyT5zKfcM49PrN0jj2Z6ZQy8MvrI40IT2Go7esrMmo');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com"+ "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "HashiCorp Terraform Enterprise Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1_-RMgUk4AwhPzXkbiZhG5KD8F2l7Gtq667xoSkdssEg');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com"+ "," +"srujana.thakkallapelly@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
     accountabilitySupport = "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "HashiCorp Vault Enterprise Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1BH6fJ4Kd2_-9dYZMFrttXn5_AO7UPIQidP4RNMRFCjs');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" ;
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "sumedha.shetty@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Pixel streaming Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1y0OWFI5RolT_X9u0dqln3b3GAEhxaZAy3oY2Wm_p0yg');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" ;
     responsibility = "meghana.mk@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

 case "Anthos Setup (For Apigee hybrid / On prem private GCP services connectivity)":
   googleSheetTemplateFile = DriveApp.getFileById('1Aqu8YG4uN-Cx9scebrFQwLJuZZ67cwraMJULRdcJrgs');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" ;
      responsibility = "meghana.mk@niveussolutions.com";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "Prisma Cloud - Palo Alto setup for AWS/ GCP/ Azure/ Other CSP":
   googleSheetTemplateFile = DriveApp.getFileById('1os5psKV-0hpq9i0VFmsizy9EkXXjV5t0w5YnM2w6SOs');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com";
     responsibility = "";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "File Share Migration":
   googleSheetTemplateFile = DriveApp.getFileById('15PJiHnp8rkTC-Ci0RRMP1ZBceEU6aKqLr-VcD3eBydc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com";
     responsibility = "";
     accountability = "suraj.rao@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "On-Prem Storage Migration to GCS":
   googleSheetTemplateFile = DriveApp.getFileById('1GJvTdr9A5EF1zVjUJnrUalA8ndAUYPaVJDFHe8Yrxjc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com";
    responsibility = "";
     accountability = "suraj.rao@niveussolutions.com";
     accountabilitySupport = "fatima.reihab@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
 
 case "NGFW - Palo Alto Creation":
   googleSheetTemplateFile = DriveApp.getFileById('18mTJri_GGLmugoqRjQ1x2MimlFkZV9cGo0OPl6tR1M4');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "vineeth.kumar@niveussolutions.com" + "," + "sumedha.shetty@niveussolutions.com";
     responsibility = "";
     accountability = "meghana.mk@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;
   
 case "Citrix based VDI Setup":
   googleSheetTemplateFile = DriveApp.getFileById('1ivxKIhQ9u2IK6VuBYDPK6qpI5N8QAVhlmtl_wVn3MOk');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," + "sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com"+ "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com";
     accountability = "suraj.rao@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

  case "SAP":
   googleSheetTemplateFile = DriveApp.getFileById('1hNwdVGtjW5CjTW11s_lNCalFd1b5clhD7Eui7nC8Az8');
   //ccRecipients="suhan.shetty@niveussolutions.com";
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "";
     accountabilitySupport = "";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

   case "CloudOps":
    if (cloudOpsOppty == 'Cloud Ops - Managed Services Support')
    {
      googleSheetTemplateFile = DriveApp.getFileById('1DmqK4sof3CzouJh1hz1JKYIgv9k9N61z3v8yogtYjzw');
    }
    if (cloudOpsOppty == 'Cloud Ops - Finops Looker Dashboard')
    {
      googleSheetTemplateFile = DriveApp.getFileById('1Ed2KwUU0_ZOozXmjdbQ83AaRGMszPpbldFryVV1i5Oc');
    }
   //ccRecipients="suhan.shetty@niveussolutions.com";
    responsibility = "prathiksha.kamath@niveussolutions.com";
     accountability = "";
     accountabilitySupport = "";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
   break;

  // case "Custom":
  //  googleSheetTemplateFile = DriveApp.getFileById('1XzDv98DcDso3cF7EOdoiGhU9KAbJ9w53XHEGAzcWH8g');
  // //  ccRecipients="suhan.shetty@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com";
  //    responsibility ="suhan.shetty@niveussolutions.com";
  //    accountability = "";
  //    accountabilitySupport = "";
  //    consult = "suhan.shetty@niveussolutions.com";
  //    informed = "";
  //    ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," + informed;
  //  break;

  default:
  googleSheetTemplateFile = DriveApp.getFileById('1ZAJEd9Fupxr-xTN2iCq-t5WVheYI2rUi4YamgLyEPhc');
   //ccRecipients="bharath.yaragatti@niveussolutions.com"+ "," +"suhan.shetty@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" + "," + "srujana.thakkallapelly@niveussolutions.com"+ "," +"sumedha.shetty@niveussolutions.com"+ "," + "vineeth.kumar@niveussolutions.com" + "," + "meghana.mk@niveussolutions.com" + "," + "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com";
     responsibility = "meghana.mk@niveussolutions.com" + "," + "prathiksha.kamath@niveussolutions.com" ;
     accountability = "suraj.rao@niveussolutions.com"+ "," +"meghana.mk@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     accountabilitySupport = "sumedha.shetty@niveussolutions.com"+ "," +"fatima.reihab@niveussolutions.com" + "," +
      "srujana.thakkallapelly@niveussolutions.com";
     consult = "suhan.shetty@niveussolutions.com";
     informed = "bharath.yaragatti@niveussolutions.com";
     ccRecipients = responsibility + "," + accountability + "," + accountabilitySupport + "," +  consult + "," +  informed;
     break;
 }
 console.log("getGoogleSheetTemplateFile ccRecipients",ccRecipients);
 return [googleSheetTemplateFile,ccRecipients];
}

//---------------------------------------------------------------------------------
var archDiagramID = '';
var effortSheetID =  '';


//Get TemplateFile for App
function getGoogleSheetTemplateFileApp(opptyType)
{
  var proposal_template = DriveApp.getFileById('175xsTPWIN25R35AXZoVLZJfp-BAOQeKfMw6hxt3J4Z4');
  console.log("opptyType",opptyType);

 if(opptyType == 'Apigee X Setup and Implementation'){
   var template = DriveApp.getFileById('10cB43xeNOCJUdtzXYnEynAGWL5Y6CGhQ0eC7FXLgMXw');
   archDiagramID = '1RM1A-9CHHfGvm_4GbIfjOIftVvhmewW';
   effortSheetID = '1eiVwUFMajQlePL1GWKV6YkqRvPablnWLK4vPIONUs-s';
 }

 else if(opptyType == 'Apigee Hybrid Setup and Implementation'){
   var template = DriveApp.getFileById('1bGBs6i6uPgtyw_4rytZ4q8T70FjAkFwF7mC5Agmaw_U');
   archDiagramID = '1SkmkvnTwc2L5S2WjFKMmzRyW5wu56PVt';
   effortSheetID = '10AlQlVHzr9DiqU2tYXTEVk9umdkPpTRU85vBDO9sM4A';
 }

 else if(opptyType == 'FlutterFlow Web + Mobile Development' || opptyType == 'FlutterFlow WebApp Development' ||      opptyType == 'FlutterFlow MobileApp Development'){
   var template = DriveApp.getFileById('1qe_MbiiVELvZqPh425WSg_hrFimrPWSZJyILQ_0_fHY');
   archDiagramID = '1vTRZdMeadwtWGczHYXym53690AepXUhN';
   effortSheetID = '1C5mjuQGSEH1CiPsIczHhVkdtu9Ozved-yW7F1lo8QoY';
 }

 else{
  var template = DriveApp.getFileById('1qe_MbiiVELvZqPh425WSg_hrFimrPWSZJyILQ_0_fHY');
  archDiagramID = '1K74Xeb92uhKlS6YeSZKzBBMqZhiPP1Y9';
  effortSheetID = '1C5mjuQGSEH1CiPsIczHhVkdtu9Ozved-yW7F1lo8QoY';
 }
  return [template, archDiagramID, effortSheetID,proposal_template];
}
//----------------------------------------------------------------------------------

//Get templateFile for Data
function getGoogleSheetTemplateFileData(opptyType)
{
  console.log("opptyType",opptyType);
  var template = DriveApp.getFileById('1wTYQoGP8Ie8HNqIkGec4lLzybRZQlpvB-nIXwODz93Q');
  return template;
}

//----------------------------------------------------------------------------------

function onFormSubmit(e)
{
 //Get latest response values with their headers

 var rowRange = getResponseSheetHeaderAndValue('Sales to CE');
 console.log("rowRange.rowValues",rowRange.rowValues);
 
//----------------------------------------------------------------------------------

 var vertical = rowRange.rowValues[2];
 console.log("Vertical",vertical);
 var leadingVertical = rowRange.rowValues[3];
 var clientName = rowRange.rowValues[4];
 clientName = convertToTitleCase(clientName);
 console.log("ClientName",clientName);

 //Destination folder
 const destinationFolder = mainDriveFunction(clientName,leadingVertical);
 console.log("destinationFolder",destinationFolder);
 var dest = DriveApp.getFolderById(destinationFolder);
 console.log("destinationFolder",dest);
 // Move additional files uploaded
 moveFiles('1cpbCjAh-T_w6LL_8p4hyCmQ_UvAqd9Jkw3wrkFvULBbhU5F6MasFvm1nzC8n5cDy-RkhQgzq',dest);
 // Move A to A costing file uploaded
 moveFiles('1jRGwibHYlZTpHDP1VwjLipE6LW4L6X5X6C7MqWhrvVld8a6ZyE0792E7myWpFQY6TGbQUQhu',dest);

  // Move the file to the designated folder
  function moveFiles(sourceFileFolder,dest) {
    var sourceFolder = DriveApp.getFolderById(sourceFileFolder); // ID of the source folder
    var files = sourceFolder.getFiles();

    while (files.hasNext()) {
      var file = files.next();
      var fileCopy = file.makeCopy(file.getName(), dest);
      dest.createFile(fileCopy.getBlob());
      DriveApp.getFileById(file.getId()).setTrashed(true); // Use setTrashed to move to trash
      console.log("File", file.getName(), "has been copied and removed from the source folder.");
      // var file = files.next();
      // dest.createFile(file.getBlob());
      // console.log("File", file);
      // sourceFolder.removeFile(file);
    }
  }


 var scopeSheetID = '';
 var infraRegistryID = '';
 scopeSheetID = '19xxhhcp1b97dp9KeLQcx08Gy07dIhfCkSWvRp8HRSyU';

 //Name the Documents
 var scopeSheetName = clientName +' - Scope Sheet Details - First Draft';
 var sowTemplateName = clientName +'- SOW - First Draft'; 
 var proposalTemplateName = clientName +'- Proposal Draft'; 
 var effortSheetName = clientName +'- Effort Sheet'; 
 var archDiagramName = clientName;
 var infraRegistrySheetName = clientName +'- Infrastructure Registry'; 
 var timelineSheetName = clientName +'- Timeline';
 var retroSheetName = clientName + '- Retro Sheet';
 var handBookName = clientName + '- CloudOps Onboarding Handbook';
 var cloudOpsTemplateName = clientName + '-CloudOps (Managed Services)';
 var finOpsTemplateName = clientName + '-Cloud FinOps Dashboard SOW';
 console.log("org sub file name",scopeSheetName);

//Opp Type 
var opptyTypeInfra = rowRange.rowValues[31];
var opptyTypeApp = rowRange.rowValues[21];
var opptyTypeData = rowRange.rowValues[35];
var opptyType = '';
//date of creation
var creationDate = todaysdate(rowRange.rowValues[0]);
console.log("Date", creationDate);
var projectOverview = rowRange.rowValues[11];
var cloudOpsOppty = rowRange.rowValues[47];

console.log("**cloudopsoppty",cloudOpsOppty)
//Choose vertical and oppty type to create deliverables

if(vertical == 'Infra' && opptyTypeInfra == 'CloudOps'){
  //var [googleSheetTemplateFile,ccRecipients] = getGoogleSheetTemplateFileInfra(opptyTypeInfra);{
    var [googleSheetTemplateFile,ccRecipients] = getGoogleSheetTemplateFileInfra(opptyTypeInfra, cloudOpsOppty);
  if(cloudOpsOppty == 'Cloud Ops - Managed Services Support'){
    var ccRecipients = 'infra-ce@niveussolutions.com' 
    createTemplateCloudOps(googleSheetTemplateFile,clientName,dest,opptyType);
    var handBookID = '1lBjoNZW3tim3SKKTqq4A5G0bXIykr6fMiglRze_s5h8'
    createHandBook(handBookID,clientName, dest,handBookName);
  }
  if(cloudOpsOppty == 'Cloud Ops - Finops Looker Dashboard')
  {
    var ccRecipients = 'infra-ce@niveussolutions.com' 
    effortSheetID = '1VTyNuC2MIDduOvBEz7Urej4h63X4fNnbR4D1f07SKXI';
    createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
    createTemplateFinOps(googleSheetTemplateFile,clientName,dest,opptyType);
  }
}
 else if(vertical == 'Infra' || vertical == 'App + Infra' || vertical == 'Infra + Data' || vertical == 'All'){
  opptyType = opptyTypeInfra;
  console.log("sss",opptyType);
  var [googleSheetTemplateFile,ccRecipients] = getGoogleSheetTemplateFileInfra(opptyTypeInfra);
  var ccRecipients = 'infra-ce@niveussolutions.com' 
  if(opptyType == 'Landing Zone Essential' || opptyType == 'Landing Zone Standard' || opptyType == 'Landing Zone Enterprise')
  {
    effortSheetID = '1k19i_WvSkyZitpsq7QWR0xJRso-ereZLRW8RsTmt11o';
  }
  else
  {
      effortSheetID = '1L1qibUUv7EtYz9D8Q4SvY4jZIv5hQbVMWzp-eOmasEo';
  }
  archDiagramID = '1QtKq9-JfR2fSsN7Wpq3pH7M68DgNqgJm';
  infraRegistryID = '1Jkgc00VELqS4HMJx94jWA62N131HPhiHR1WJ-4LOC2c';
  timelineID = '156qKqPgIitJsO_113DWxvky6n4PI9RZgcQPTzRzRW54';
  retroSheetID = '1rlfQdWvuqFKzhucmRzIDzI9mPzPj215W37yiA-Zm8bk'
  createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
  createSOW(googleSheetTemplateFile,clientName,dest,opptyType);
  createInfraRegistry(infraRegistryID,dest,infraRegistrySheetName);
  createTimelineSheet(timelineID,dest,timelineSheetName);
  createRetroSheet(retroSheetID,dest,retroSheetName);
}
 if(vertical == 'Data' || vertical == 'App + Data' || vertical == 'Infra + Data' || vertical == 'All'){
  opptyType = opptyTypeData;
  var googleSheetTemplateFile = getGoogleSheetTemplateFileData(opptyTypeData);
  var ccRecipients = 'ce-data@niveussolutions.com' + "," + 'ce-partnerBU@niveussolutions.com';
  effortSheetID = '1zGxP9-tQlVMo7ejHQyjCt8JQxBow6xJNWjvlqgBNcGg';
  archDiagramID = '1QtKq9-JfR2fSsN7Wpq3pH7M68DgNqgJm';
  //effortSheetID = '1EOizLsYiHidH3Vesm4c9evWvcC8_NNViL16mY5m9ywo'
  createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
  createSOW(googleSheetTemplateFile,clientName,dest,opptyType);
 }
  if(vertical == 'App' || vertical == 'App + Infra' || vertical == 'App + Data' || vertical == 'All'){
  opptyType = opptyTypeApp;
  var [googleSheetTemplateFile,archDiagramID, effortSheetID,googleProposalFile] = getGoogleSheetTemplateFileApp(opptyTypeApp);
  var ccRecipients = 'ce-app-mod@niveussolutions.com';
  createEffortSheet(effortSheetID,clientName,dest,effortSheetName);
  createSOW(googleSheetTemplateFile,clientName,dest,opptyType);
  createProposalDraft(googleProposalFile,clientName,dest,opptyType);
 }

//Create Scope Sheet
 scopeSheetCopy = createScopeSheet(scopeSheetID,dest,scopeSheetName);
  var newscopeSheetID = scopeSheetCopy.getId();
  updateScopeSheetValue(rowRange.rowHeaderValue,rowRange.rowValues,newscopeSheetID);

if(vertical == 'App + Infra'){
  ccRecipients = 'infra-ce@niveussolutions.com' + "," + 'ce-app-mod@niveussolutions.com';
  console.log("ccReceipient 1234 ", ccRecipients);
}
else if(vertical == 'App + Data'){
  ccRecipients = 'ce-data@niveussolutions.com' + "," + 'ce-app-mod@niveussolutions.com';
  console.log("ccReceipient 1234 ", ccRecipients);
}
else if(vertical == 'Infra + Data'){
  ccRecipients = 'ce-data@niveussolutions.com' + "," + 'infra-ce@niveussolutions.com';
}
else if(vertical == 'All'){
  ccRecipients = 'ce-data@niveussolutions.com' + "," + 'infra-ce@niveussolutions.com' + "," +  'ce-app-mod@niveussolutions.com';
}

 
//----------------------------------------------------------------------------------


 //Get SOW
 function createSOW(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(sowTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
 }

 function createProposalDraft(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(proposalTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
 }
 //get Spreadsheet 
function createScopeSheet(scopeSheetID,dest,scopeSheetName){
  const scopeSheet = DriveApp.getFileById(scopeSheetID);
  const scopeSheetCopy = scopeSheet.makeCopy(scopeSheetName, dest);
  return scopeSheetCopy;
  }
//get Effort Sheet
function createEffortSheet(effortSheetID,clientName,dest){
  const effortSheet = DriveApp.getFileById(effortSheetID);
  const effortSheetCopy = effortSheet.makeCopy(effortSheetName, dest);
  const newEffortSheetID = effortSheetCopy.getId();
  updateEffortSheetValues(clientName,newEffortSheetID);
  }
  //Create Infra Registry Sheet
function createInfraRegistry(infraRegistryID,dest,infraRegistrySheetName){
   const infraRegistrySheet = DriveApp.getFileById(infraRegistryID);
   const effortSheetCopy = infraRegistrySheet.makeCopy(infraRegistrySheetName, dest);
}
//Create Timeline Sheet
function createTimelineSheet(timelineID,dest,timelineSheetName){
   const timelineSheet = DriveApp.getFileById(timelineID);
  const timelineSheetCopy = timelineSheet.makeCopy(timelineSheetName, dest);
}

//Create Retro Sheet
function createRetroSheet(retroSheetID,dest,retroSheetName){
  const retroSheet = DriveApp.getFileById(retroSheetID);
  const retroSheetCopy = retroSheet.makeCopy(retroSheetName, dest);
}

// Cloud Ops Template
function createTemplateCloudOps(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(cloudOpsTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
}
//Finops Template
function createTemplateFinOps(googleSheetTemplateFile,clientName,dest,opptyType){
 var id = googleSheetTemplateFile.getId();
 const newId = DriveApp.getFileById(id);
 const sowTemplate = newId.makeCopy(finOpsTemplateName, dest);
 const sowTemplateID = sowTemplate.getId();
 console.log("sowTemplateID",sowTemplateID);
 updateSowValues(clientName,sowTemplateID,opptyType);
}
//Create Handbook
function createHandBook(handBookID,clientName, dest,handBookName){
  console.log("HandbookID",handBookID);
  const handBook = DriveApp.getFileById(handBookID);
  const handBookCopy = handBook.makeCopy(handBookName, dest);
  const newhandBookID = handBookCopy.getId();
  updateHandBookValue(clientName,newhandBookID);
}

 //Arch Diagram
if(opptyTypeInfra != 'CloudOps'){
  console.log("Opp",opptyTypeInfra);
  var archDiagram = DriveApp.getFileById(archDiagramID);
  var archDiagramCopy = archDiagram.makeCopy(archDiagramName, dest);
}
//Update Scope Sheet Value
function updateScopeSheetValue(responseHeaderValues,responseValues,sheetID){
 var ss = SpreadsheetApp.openById(sheetID);
 SpreadsheetApp.setActiveSpreadsheet(ss);
 ss.getRange('B1').setValue(responseValues[0].toString());
 for (var i=1;i<=responseHeaderValues.length;i++){
   //Setup the Response headers
   ss.getRange('A'+ i.toString()).setValue(responseHeaderValues[i-1].toString());
   //Setup the Response ValuescopiedSheetID
   ss.getRange('B'+ i.toString()).setValue(responseValues[i-1].toString());
 }
}
//Update SOW Values
function updateSowValues(clientName,docID,opptyType){
  const doc = DocumentApp.openById(docID);
  const body = doc.getBody();
  body.replaceText('{{Client Name}}',clientName);
  body.replaceText('{{Date of SOW}}',creationDate);
  body.replaceText('{{Project Overview}}',projectOverview);
  body.replaceText('{{Opp Type}}',opptyType);
}

//Update Effort Sheet Values
function updateEffortSheetValues(clientName,sheetID){
  const sheet = SpreadsheetApp.openById(sheetID);
  var textFinder = sheet.createTextFinder("{{Client Name}}");
  textFinder.replaceAllWith(clientName);
}

//Update Handbook Values
function updateHandBookValue(clientName,handBookID){
  const doc = DocumentApp.openById(handBookID);
  const body = doc.getBody();
  body.replaceText('{{Client Name}}',clientName);
  body.replaceText('{{Date of SOW}}',creationDate);
}

 //var recipient_email = rowRange.rowValues[1].toString();
 //sendEmail(org_name, copiedSheetID,recipient_email,mailID,ccRecipients);
 
// const oppType = rowRange.rowValues[31];

 const typeOpp = rowRange.rowValues[8];
 const salesSpocEmail= rowRange.rowValues[1];
 salesSpocName =extractNameFromEmail(salesSpocEmail);
 console.log("SalesSpoc",salesSpocName);
 const hubSpotLink = rowRange.rowValues[6];
 const dealSize = rowRange.rowValues[15];
 var constraints = rowRange.rowValues[20];
 console.log("** Get Date",rowRange.rowValues[18]);
 const nextCall = toDateFmt(rowRange.rowValues[18]);
 console.log("formattedDate = ",nextCall);
 //const problemStatement = rowRange.rowValues[9];
 if(constraints === "")
 constraints = 'Not Applicable';
 console.log("Constraints",constraints);
 var link = 'https://drive.google.com/drive/folders/'+destinationFolder;
 if(vertical == 'All')
 vertical = 'App + Infra + Data';
 ccRecipients = ccRecipients + "," +salesSpocEmail + "," + "salesops@niveussolutions.com" + "," + "suyog.shetty@niveussolutions.com";
 console.log(ccRecipients);
 //ccRecipients = 'fatima.reihab@niveussolutions.com';
   sendEmail(opptyType,vertical,clientName,salesSpocName,typeOpp,projectOverview,hubSpotLink,dealSize,constraints,nextCall,link,ccRecipients);
}

//----------------------------------------------------------------------------------


//Step7 : get the rowValues from the Re+ "," +sponse sheet
function getResponseSheetHeaderAndValue(sheetName){
 const sheet2 = SpreadsheetApp.getActiveSpreadsheet();
 var last_row = sheet2.getLastRow() - 1;
 var range = sheet2.getDataRange();
 var rowHeaderValue = range.getValues()[0];
 var rowValues = range.getValues()[last_row];
 return {rowHeaderValue , rowValues};
}
 
//Step8 : Append the rowValues from the Response sheet to the Scope template sheet


//Sending Email
function sendEmail(typeOfDeal,vertical,client,sales,dealType,projectOverview,hubspotLink,dealSize,constraints,nextCallDate,link,ccRecipients) {
  var data ={typeOfDeal,vertical,client,sales,dealType,projectOverview,hubspotLink,dealSize,constraints,nextCallDate,link};
  var templ = HtmlService.createTemplateFromFile('templateEmail');
  templ.data=data;

  var message = templ.evaluate().getContent();
  console.log("ccReceipient",ccRecipients);
    MailApp.sendEmail({
  //Sending mail based on the verticals
    to: ccRecipients,
    //to: "fatima.reihab@niveussolutions.com" + "," + "suraj.rao@niveussolutions.com",
    subject: "New enagagement with "+client,
    htmlBody: message,
    cc: 'bharath.yaragatti@niveussolutions.com' 
  });
}

var folder = "NewFolder";

//------------------------------------------------------------------------------------------------------------
//Destination folder
function mainDriveFunction(clientName,leadingVertical) {
  if(leadingVertical == 'App Mod')
    var parentFolderID = '1tX3Q9NuPRZsvEhd2JeArRYNMiXsn_t6A'; 
  else if(leadingVertical == 'Infra Mod')
    var parentFolderID = '1ZIs7pRV8SUb2sO9yyFWXGNwL3UC2cEED';
  else if(leadingVertical == 'Data Mod')
    var parentFolderID = '1tUu9_eFLR56LqF89NEH-W8Mw4XdybOw0';
  else
    console.log("Invalid Vertical");

  const month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  const d = new Date();
  let currMonth = month[d.getMonth()];
  let currYear = d.getFullYear();
  console.log('Year is', currYear);
  var folderID = getFolderByName(parentFolderID,currYear); // Year Folder
  console.log("Folder ID is ", folderID);
  if(folderID==null) {
    createFolder(currYear, parentFolderID); // Created new folder
    folderID = getFolderByName(parentFolderID,currYear);   // Get folder ID of new folder that was created
  }
  parentFolderID = folderID; 
  console.log("Year done");

  console.log('Month is', currMonth);
  folderID = getFolderByName(parentFolderID,currMonth);      // Month Folder
  if(folderID==null) {
    createFolder(currMonth, parentFolderID); // Created new folder
    folderID = getFolderByName(parentFolderID,currMonth);   // Get folder ID of new folder that was created
  }
  parentFolderID = folderID;
  console.log("Month done");

  console.log('clientName is', clientName);
  folderID = getFolderByName(parentFolderID,clientName);      // Client's Name Folder
  if(folderID==null) {
    createFolder(clientName, parentFolderID); // Created new folder
    folderID = getFolderByName(parentFolderID,clientName);   // Get folder ID of new folder that was created
  }
  parentFolderID = folderID;
  console.log("Client done");
  console.log("parentFolderID",parentFolderID);
  Utilities.sleep(5000);
  // if(parentFolderID == null)
  //   {
  //     parentFolderIdNew = parentFolderID;
  //     var i = 3;
  //     while(parentFolderIdNew == null && i>0)
  //     {
  //       Utilities.sleep(3000);
  //       i--;
  //       var parentFolderIdNew1 = checkNullFolder(parentFolderID);
  //       console.log("parentID",parentFolderIdNew1);
  //     }
  //   }
  var folderID = getFolderByName(parentFolderID,clientName); 
  console.log("Client folderID =",folderID);
  return parentFolderID;
  //Copy SOW and other files
  //createDeliverables(folderID);
}

// Create new folder if it doesn't exist
function createFolder(folderName, folderDestination) {
  var parentFolder = DriveApp.getFolderById(folderDestination);
  var folderName = parentFolder.createFolder(folderName);
  console.log("Created Folder - "+folderName+" Folder Destination - "+folderDestination);
}
//mainDriveFunction

// Get folder id by name
function getFolderByName(parentFolderID,folderName) {
  console.log("Parent Folder",parentFolderID);
  console.log("Folder Name", folderName);
  //sometimes error here
  var folders = DriveApp.getFolderById(parentFolderID).getFoldersByName(folderName);
  var folderID = null;
  if (folders.hasNext())
    folderID = folders.next();
  if(folderID != null) {
    var folderIDdestination=folderID.getId();
    console.log("Folder Found - "+folderID+" Folder ID is - "+folderIDdestination);
    return folderID.getId();
  }
  else {
    console.log('Null Folder');
    return null;
  }
}


function toDateFmt(dt_string) {
  console.log("toDateFmt",dt_string);

  var millis = Date.parse(dt_string);
  var date = new Date(millis);

  console.log("extracted date 1",date);
  console.log("extracted date 2",date.getDate());
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  // Return the mainDriveFunctione date in dd-mm-YYYY format
  return `${day}-${month}-${year}`;
}

//Extracting sales spoc name
function extractNameFromEmail(salesSpoc) {
  var name = salesSpoc.match(/^([^@]*)@/)[1]; // Extract the name before the '@' symbol
  name = name.replace('.', ' '); // Replace dots with spaces
  name = name.replace(/\b\w/g, function(l){ return l.toUpperCase() }); // Capitalize the first letter of each word
  return name;
}

//Extracting Date
function todaysdate(timestampValue){
  console.log("Date",timestampValue)
  var date = new Date(timestampValue); // Convert the timestamp to a date object
  var dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-YYYY"); // Format the date object into a string with only the date portion
  return dateString;
}

// Title Case 
function convertToTitleCase(str) {
  const newStr = str.toLowerCase().split(' ')
    .map(w => {
      if (w) {
        return w.charAt(0).toUpperCase() + w.slice(1);
      }
      return '';
    })
    .join(' ');
  console.log(newStr);
  return newStr;
}


//Check if the Parent Folder is Null

function checkNullFolder(parentFolderID)
{
  return parentFolderID;
}
