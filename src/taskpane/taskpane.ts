/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

//fluent-ui imports
import {
  fluentTreeView,
  fluentTreeItem,
  fluentCheckbox,
  fluentSelect,
  fluentOption,
  fluentButton,
  fluentTextField,
  provideFluentDesignSystem,
} from "@fluentui/web-components";

provideFluentDesignSystem().register(
  fluentTreeView(),
  fluentTreeItem(),
  fluentCheckbox(),
  fluentSelect(),
  fluentOption(),
  fluentButton(),
  fluentTextField()
);
import { creds } from "./creds";
import jwt_decode from "jwt-decode";
//authorization token
var token = "";
var chosenPractice = "";
var indexOfPractice = 0;
var totalNumberOfPractices = 0;
var searchedForPractice = false;
var data = "";
var codicepratica = "";

//initialize
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    debugger;
    // getUserData();
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
      debugger;
      const accessToken = result.value;

      // Use the access token.
    });
    document.getElementById("CloseForm").onclick = CloseForm;
    var ClientSecretKey = Office.context.roamingSettings.get("ClientSecretKey");
    if (ClientSecretKey != undefined) {
      document.getElementById("choosePracticeButton").onclick = searchForPractice;
      document.getElementById("nextPageOfPractices").onclick = goToNextPage;
      document.getElementById("prevPageOfPractices").onclick = goToPreviousPage;
      document.getElementById("FascicoloAllegati").onclick = FascicoloAllegati;

      $.ajax({
        url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/oauth/token",
        method: "POST",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
        },
        data: {
          grant_type: "client_credentials",
          client_id: creds.client_id, //Provide your client_id
          client_secret: creds.client_secret, //Provide your client_secret
        },
        success: function (response) {
          token = response.access_token;
          //GetAllPratiche(token);
        },
        error: function (error) {
          document.getElementById("error").innerText = error.toString();
        },
      });
      $.ajax({
        crossDomain: true,
        url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/archive?limit=10", // Pass your tenant name instead of sharepointtechie
        method: "GET",
        headers: {
          "content-type": "application/x-www-form-urlencoded",
          Authorization: "Bearer " + ClientSecretKey,
        },

        success: function () {
          token = ClientSecretKey;
          GetAllPratiche(ClientSecretKey);
        },
        error: function () {
          document.getElementById("SaveToken").onclick = SaveToken;

          document.getElementById("TokenRequired").style.display = "";
          document.getElementById("startingLoader").style.display = "none";
        },
      });
    } else {
      document.getElementById("SaveToken").onclick = SaveToken;
      document.getElementById("startingLoader").style.display = "none";
      document.getElementById("TokenRequired").style.display = "";
    }
  }
});
// async function getUserData() {
//   debugger;

//   let userTokenEncoded = await OfficeRuntime.auth
//     .getAccessToken({
//       allowSignInPrompt: true,
//       forMSGraphAccess: true,
//     })
//     .then((token: any) => {
//       debugger;
//       console.log("bbbbbbbbbbbbbb");
//       var tokeni = token;
//       var url =
//         "https://graph.microsoft.com/v1.0/me/messages/AQMkAGRmMjc1MDFlLWY4NWItNGMxYS04Yjk1LTRhMjAyM2Q4MmI1MQBGAAADq6f5QYdXIUa38XX8J1ml3wcAMB8YFZIaAUucR37vvHv5-wAAAgEMAAAAMB8YFZIaAUucR37vvHv5-wAFSCGhegAAAA==";
//       var settings = {
//         url: "https://howling-crypt-47129.herokuapp.com/https://graph.microsoft.com/v1.0/me",
//         method: "GET",
//         timeout: 0,
//         headers: {
//           Authorization: "Bearer " + tokeni,
//           contentType: "application/x-www-form-urlencoded",
//         },
//         data: {
//           scope: "https://graph.microsoft.com/.default",
//         },
//       };

//       $.ajax(settings)
//         .done(function (response) {
//           debugger;
//           var test = response;
//         })
//         .catch((error: any) => {
//           debugger;
//           var error = error;
//         });
//     })
//     .catch((error: any) => {
//       debugger;
//       console.log("cccccccccccccccccc");
//     });
// }

//close form
export function CloseForm() {
  document.getElementById("myForm").style.display = "none";
}

export function FascicoloAllegati() {
  debugger;
  document.getElementById("FascicoloAllegati").style.display = "none";
  for (let i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
    var AllegatiFascicolo = document.createElement("div");
    AllegatiFascicolo.setAttribute("id", Office.context.mailbox.item.attachments[i].id);
    var obj = Office.context.mailbox.item.attachments[i].name.split(".");
    AllegatiFascicolo.innerHTML =
      `<div class="Allegati">
  <input type="text" id="IdA` +
      i +
      `" value=` +
      obj[0] +
      `><textbox disabled id="Id1` +
      i +
      `">` +
      "." +
      obj[1] +
      `</textbox><br><button id=` +
      "Allegati" +
      Office.context.mailbox.item.attachments[i].id +
      `>Fascicolo Allegati</button></div><br/>`;
    document.getElementById("ReadMessage").append(AllegatiFascicolo);
    var FascicoloAllegati = document.getElementById("Allegati" + Office.context.mailbox.item.attachments[i].id);
    FascicoloAllegati.onclick = function setChoosenPractice() {
      debugger;
      //  var test=document.getElementById("Id" + Office.context.mailbox.item.attachments[i].id).value;
      // var test= document.getElementById("Id" + Office.context.mailbox.item.attachments[i].id).innerText;
      var NameOfFile = $("#IdA" + i).val();
      var FullNameOfFile = NameOfFile + "." + Office.context.mailbox.item.attachments[i].name.split(".")[1];
      var options = { asyncContext: { type: Office.context.mailbox.item.attachments[i].attachmentType } };
      Office.context.mailbox.item.getAttachmentContentAsync(
        Office.context.mailbox.item.attachments[i].id,
        options,
        function (result) {
          if (result.status == Office.AsyncResultStatus.Succeeded) {
            debugger;
            var AttachmentContent = result.value;
            var ContentType = Office.context.mailbox.item.attachments[i].contentType;
            const url = "data:" + ContentType + ";" + "base64," + AttachmentContent.content;
            fetch(url)
              .then((res) => res.blob())
              .then((blob) => {
                debugger;
                const file = new File([blob], FullNameOfFile, { type: ContentType });
                console.log(file);
                let data = new FormData();
                data.append("file", file);
                data.append("titolodocumento", NameOfFile.toString());
                data.append("nomefile", FullNameOfFile.toString());
                data.append("data", "2020-01-01");
                data.append("codicepratica", codicepratica);

                fetch("https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/document", {
                  method: "POST",
                  body: data,
                  headers: {
                    Authorization: "Bearer " + token,
                  },
                })
                  .then(function (serverPromise) {
                    debugger;
                    serverPromise
                      .json()
                      .then(function (j) {
                        // console.log(j);
                        document.getElementById("IdA" + i).style.display = "none";
                        document.getElementById("Id1" + i).style.display = "none";
                        document.getElementById(
                          "Allegati" + Office.context.mailbox.item.attachments[i].id
                        ).style.display = "none";
                      })
                      .catch(function (e) {
                        console.log(e);
                      });
                  })
                  .catch(function (e) {
                    debugger;
                    console.log(e);
                  });
              });
          }
        }
      );
    };
  }
}

//save token
export function SaveToken() {
  debugger;
  var Token = $("#token").val();
  Office.context.roamingSettings.set("ClientSecretKey", Token);
  $.ajax({
    crossDomain: true,
    url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/archive?limit=10", // Pass your tenant name instead of sharepointtechie
    method: "GET",
    headers: {
      "content-type": "application/x-www-form-urlencoded",
      Authorization: "Bearer " + Token,
    },

    success: function () {
      document.getElementById("Messagio").style.display = "";
      document.getElementById("SaveToken").style.display = "none";
      document.getElementById("token").style.display = "none";
    },
    error: function () {
      document.getElementById("SaveToken").onclick = SaveToken;
      document.getElementById("TokenRequired").style.display = "";
      document.getElementById("startingLoader").style.display = "none";
    },
  });
  Office.context.roamingSettings.saveAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
    } else {
      document.getElementById("TokenRequired").style.display = "none";
      document.getElementById("Messagio").style.display = "";
      document.getElementById("SaveToken").style.display = "none";
      document.getElementById("token").style.display = "none";
      Office.onReady();
    }
  });
}

//get all practices
export async function GetAllPratiche(token) {
  var settings = {
    url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/archive?intestazione=true",
    method: "GET",
    timeout: 0,
    headers: {
      Authorization: "Bearer " + token,
    },
  };

  await $.ajax(settings).done(function (response) {
    var practices = response.payload
      .sort((a, b) => {
        return a.data - b.data;
      })
      .reverse();
    totalNumberOfPractices = practices.length;
    practices.slice(indexOfPractice, indexOfPractice + 10).forEach((item) => {
      var Inestazione = "Intestazione pratica";
      Inestazione = item.intestazione;
      var practiceButton = document.createElement("div");
      practiceButton.setAttribute("id", item.id);
      practiceButton.innerHTML =
        `<div class="practices">
        <span>` +
        item.id +
        ` - ` +
        item.avvocato +
        " - " +
        Inestazione +
        `</span><br><button id=` +
        "Dettagli" +
        item.id +
        `>Dettagli</button><button id=` +
        "Scegli" +
        item.id +
        `>Scegli Pratica</button></div><br/>`;
      document.getElementById("practices").append(practiceButton);
      var ScegliPratica = document.getElementById("Scegli" + item.id);
      ScegliPratica.onclick = function setChoosenPractice() {
        chosenPractice = item.id;
        data = item.data;
        codicepratica = item.codicearchivio;
        if (Office.context.mailbox.item.displayReplyForm != undefined) {
          document.getElementById("ReadMessage").style.display = "";
          document.getElementById("practicesSection").style.display = "none";
        } else {
          chooseTemplateOrDocuments(); //------------------------------------------------------------------------------------------------------------> Here here - new mail
        }
      };
      var Dettagli = document.getElementById("Dettagli" + item.id);
      Dettagli.onclick = function setChoosenPractice() {
        document.getElementById("IDForm").setAttribute("value", item.id);
        document.getElementById("AvvocatoForm").setAttribute("value", item.avvocato);
        document
          .getElementById("RuologeneraleForm")
          .setAttribute("value", item.ruologeneraleanno + "/" + item.ruologeneralenumero);
        document.getElementById("DataForm").setAttribute("value", item.data);
        document.getElementById("CodicearchivioForm").setAttribute("value", item.codicearchivio);
        document.getElementById("DescrizioneForm").setAttribute("value", item.descrizione);
        document.getElementById("StatoForm").setAttribute("value", item.stato);
        document.getElementById("myForm").style.display = "block";
      };
      var modal = document.getElementById("myModal");
      window.onclick = function (event) {
        if (event.target == modal) {
          modal.style.display = "none";
        }
      };
    });
  });
  document.getElementById("startingLoader").style.display = "none";
  document.getElementById("practicesLoader").style.display = "none";
  document.getElementById("practicesSection").style.display = "flex";
  if (indexOfPractice > totalNumberOfPractices - 10) {
    document.getElementById("nextPageOfPractices").setAttribute("disabled", "");
  } else {
    document.getElementById("nextPageOfPractices").removeAttribute("disabled");
  }
}

//go to next page of practices
export async function goToNextPage() {
  document.getElementById("prevPageOfPractices").removeAttribute("disabled");
  document.getElementById("practices").innerHTML = "";
  document.getElementById("practicesLoader").style.display = "";
  document.getElementById("nextPageOfPractices").setAttribute("disabled", "");
  indexOfPractice = indexOfPractice + 10;
  GetAllPratiche(token);
}

//go to previous page of practices
export async function goToPreviousPage() {
  document.getElementById("prevPageOfPractices").removeAttribute("disabled");
  document.getElementById("practices").innerHTML = "";
  document.getElementById("practicesLoader").style.display = "";
  document.getElementById("nextPageOfPractices").setAttribute("disabled", "");
  indexOfPractice = indexOfPractice - 10;
  if (indexOfPractice == 0) {
    document.getElementById("prevPageOfPractices").setAttribute("disabled", "");
  }
  GetAllPratiche(token);
}

//search for practice section
export async function searchForPractice() {
  document.getElementById("selectedPractices").style.display = "none";
  document.getElementById("searchNoResults").style.display = "none";
  searchedForPractice = true;
  //get all users
  var settings = {
    url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/user",
    method: "GET",
    timeout: 0,
    headers: {
      Authorization: "Bearer " + token,
    },
  };

  await $.ajax(settings).done(function (response) {
    var users = response.payload;
    users.sort().forEach((user) => {
      var option = document.createElement("fluent-option");
      option.setAttribute("id", user.id);
      option.setAttribute("value", user.id);
      option.innerText = user.nomeutente;
      document.getElementById("userList").append(option);
    });
  });
  //display searchPractice section
  document.getElementById("searchPractices").style.display = "table";
  document.getElementById("practicesSection").style.display = "none";
  //go back to all practices button
  document.getElementById("goBackToSelectPractice").onclick = function () {
    document.getElementById("searchPractices").style.display = "none";
    document.getElementById("startingLoader").style.display = "";
    GetAllPratiche(token);
  };

  //search for practice by set parameters
  document.getElementById("submitPracticeSearch").onclick = async function getPractices() {
    debugger;
    var paramString;
    var userId = document
      .querySelector('fluent-option[aria-selected="true"]')
      .getAttribute("id")
      .replace("option-", "");
    var searchBy = document.getElementById("ricerca_testo")["value"];
    var beginDate = document.getElementById("beginDate")["value"];
    var endDate = document.getElementById("endDate")["value"];

    var searchByString = "ricerca_testo=" + searchBy;
    var userIdString = "id_utente_soggetto=" + userId;
    var beginDateString = "data_dal" + beginDate + " 00:00:00";
    var endDateString = "data_al" + endDate + " 00:00:00";
    if (beginDate == null || beginDate == "") {
      beginDateString = "";
    }

    if (endDate == null || endDate == "") {
      endDateString = "";
    }

    if (searchBy == null || searchBy == "") {
      searchByString = "";
    }

    if (userId == null || userId == "") {
      userIdString = "";
    }
    //paramString
    paramString = "?" + searchByString + "&" + userIdString + "&" + beginDateString + "&" + endDateString;
    var settings = {
      url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/archive" + paramString,
      method: "GET",
      timeout: 0,
      headers: {
        Authorization: "Bearer " + token,
      },
    };
    await $.ajax(settings).done(function (response) {
      var practices = response.payload;
      if (!response.payload[0]) {
        document.getElementById("searchNoResults").style.display = "";
      }
      practices.forEach((item) => {
        var Inestazione = "Intestazione pratica";
        var practiceButton = document.createElement("div");
        practiceButton.setAttribute("id", item.id);
        practiceButton.setAttribute("appearance", "stealth");
        practiceButton.innerHTML =
          `<div class="practices">
          <span>` +
          item.id +
          ` - ` +
          item.avvocato +
          " - " +
          Inestazione +
          `</span><br><button id=` +
          "DettagliSearch" +
          item.id +
          `>Dettagli</button><button id=` +
          "ScegliSearch" +
          item.id +
          `>Scegli Pratica</button></div><br/>`;
        document.getElementById("selectedPractices").append(practiceButton);
        var ScegliPratica = document.getElementById("ScegliSearch" + item.id);
        ScegliPratica.onclick = function setChoosenPractice() {
          chosenPractice = item.id;

          if (Office.context.mailbox.item.displayReplyForm != undefined) {
            //------------------------------------------------------------------------------------------------------------------------------------------> Here here - received mail
          } else {
            chooseTemplateOrDocuments(); //-------------------------------------------------------------------------------------------------------------> Here here - new mail
          }
        };
        var Dettagli = document.getElementById("DettagliSearch" + item.id);
        Dettagli.onclick = function setChoosenPractice() {
          document.getElementById("IDForm").setAttribute("value", item.id);
          document.getElementById("AvvocatoForm").setAttribute("value", item.avvocato);
          document
            .getElementById("RuologeneraleForm")
            .setAttribute("value", item.ruologeneraleanno + "/" + item.ruologeneralenumero);
          document.getElementById("DataForm").setAttribute("value", item.data);
          document.getElementById("CodicearchivioForm").setAttribute("value", item.codicearchivio);
          document.getElementById("DescrizioneForm").setAttribute("value", item.descrizione);
          document.getElementById("StatoForm").setAttribute("value", item.stato);
          document.getElementById("myForm").style.display = "block";
        };
        var modal = document.getElementById("myModal");
        window.onclick = function (event) {
          if (event.target == modal) {
            modal.style.display = "none";
          }
        };
      });
    });
    document.getElementById("goBackToSearch").onclick = searchForPractice;
    document.getElementById("selectedPractices").style.display = "flex";
    document.getElementById("searchPractices").style.display = "none";
  };
}

//choose template
export async function chooseTemplate(id) {
  //request to get all templates for the specific practice id
  // document.getElementById("displayChosenPracticeTemplate").innerText = "Practice Id = " + id;
  document.getElementById("practicesSection").style.display = "none";
  document.getElementById("selectedPractices").style.display = "none";
  id = "26";
  var settings = {
    url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/emails-storage?file_id=" + id,
    method: "GET",
    timeout: 0,
    headers: {
      Authorization: "Bearer " + token,
    },
  };

  var option;
  //create the template options
  await $.ajax(settings).done(function (response) {
    var templates = response.payload;
    templates.forEach((template) => {
      option = document.createElement("fluent-option");
      option.setAttribute("class", "templateOption");
      option.setAttribute("id", template.id);
      option.innerText = template.title;
      document.getElementById("selectTemplate").append(option);
    });
  });
  // chooseDocument(chosenPractice);
  document.getElementById("chooseTemplateOrDocs").style.display = "none";
  document.getElementById("searchPractices").style.display = "none";
  document.getElementById("templateChooser").style.display = "flex";
  // document.getElementById("documentChooser").style.display = "flex";
  document.getElementById("addTemplate").onclick = addTemplate;
  document.getElementById("goBackToTempDocChoiceTemp").onclick = chooseTemplateOrDocuments;
  return false;
}

//choose between template or documents
export async function chooseTemplateOrDocuments() {
  document.getElementById("documents").innerHTML = "";
  document.getElementById("selectTemplate").innerHTML = "";
  document.getElementById("templateChooser").style.display = "none";
  document.getElementById("documentChooser").style.display = "none";
  document.getElementById("practicesSection").style.display = "none";
  document.getElementById("selectedPractices").style.display = "none";
  document.getElementById("chooseTemplateOrDocs").style.display = "";
  document.getElementById("goBackToPractice_ChosenPractice").onclick = function () {
    if (searchedForPractice == true) {
      document.getElementById("selectedPractices").style.display = "";
      document.getElementById("chooseTemplateOrDocs").style.display = "none";
    } else {
      document.getElementById("practicesSection").style.display = "";
      document.getElementById("chooseTemplateOrDocs").style.display = "none";
    }
  };
  document.getElementById("chooseTemplate").onclick = function () {
    chooseTemplate(26);
  };
  document.getElementById("chooseDocuments").onclick = function () {
    chooseDocument(chosenPractice);
  };
}

//add selected template
export async function addTemplate() {
  var templateId = document.querySelector('fluent-option[aria-selected="true"]').getAttribute("id");
  var body;
  var settings = {
    url:
      "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/emails-storage?id=" + templateId,
    method: "GET",
    timeout: 0,
    headers: {
      Authorization: "Bearer " + token,
    },
  };

  await $.ajax(settings).done(function (response) {
    body = response.payload[0].body;
    Office.context.mailbox.item.body.setSelectedDataAsync(body, {
      coercionType: Office.CoercionType.Html,
      asyncContext: { var3: 1, var4: 2 },
    });
  });

  //choose documents to add
  // chooseDocument(chosenPractice);
  return false;
}

export async function PostDocument(file, NomeFile) {
  debugger;
  var paramString =
    "?file=" +
    file +
    "&titolodocumento=" +
    NomeFile +
    "&nomefile=" +
    NomeFile +
    "&data=" +
    data +
    "&codicepratica=" +
    codicepratica;
  var settings = {
    url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/document" + paramString,
    method: "POST",
    timeout: 0,
    headers: {
      Authorization: "Bearer " + token,
    },
  };
  await $.ajax(settings)
    .done(function (response) {
      debugger;
      var Response = response;
    })
    .fail(function (jqXHR, textStatus) {
      debugger;
      var test = textStatus;
      var newtest = jqXHR;
    });
}

//documents section
export async function chooseDocument(id) {
  document.getElementById("chooseTemplateOrDocs").style.display = "none";
  // document.getElementById("displayChosenPracticeDocument").innerText = "Practice Id = " + id;
  document.getElementById("documentChooser").style.display = "flex";
  //setting id=2 static for testing
  // id = 2; //should be removed
  var settings = {
    url:
      "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/file-documents-tree?codicepratica=" +
      id,
    method: "GET",
    timeout: 0,
    headers: {
      Authorization: "Bearer " + token,
    },
  };
  document.getElementById("goBackToTempDocChoiceDoc").onclick = chooseTemplateOrDocuments;
  document.getElementById("documents").style.display = "none";
  //create tree-view from documentTreeView request
  var elementId;
  document.getElementById("templateChooser").style.display = "none";
  await $.ajax(settings).done(function (response) {
    response.payload[0].nodes.forEach((node) => {
      var parent = document.createElement("fluent-tree-item");
      var parentNodeCheckbox = document.createElement("fluent-checkbox");
      var parentNode = document.createElement("span");
      parentNode.innerText = node.name;
      elementId = "nodeId" + node.id;
      parent.setAttribute("id", elementId);
      parent.append(parentNodeCheckbox);
      parent.append(parentNode);
      if (node.type == "folder") {
        elementId = "nodeId" + node.id;
        parent.setAttribute("class", "parentFolder");
        node.nodes.forEach((item) => {
          var childParent = document.createElement("fluent-tree-item");
          var childNodeCheckbox = document.createElement("fluent-checkbox");
          var childNodeItem = document.createElement("span");
          childNodeItem.innerText = item.name;
          childParent.append(childNodeCheckbox);
          childParent.append(childNodeItem);
          childParent.setAttribute("id", "735");
          childParent.setAttribute("class", "childFile");
          parent.append(childParent);
        });
        document.getElementById("documents").append(parent);
      } else {
        elementId = node.id;
      }
      parent.setAttribute("id", elementId);
      document.getElementById("documents").append(parent);
    });
  });
  document
    .querySelectorAll("fluent-tree-view > fluent-tree-item > fluent-tree-item, fluent-tree-view > fluent-tree-item")
    .forEach((item) => {
      item.shadowRoot.querySelector("slot").innerHTML =
        '<svg xmlns="http://www.w3.org/2000/svg" fill="#000000" viewBox="0 0 24 24" width="24px" height="24px"><path d="M 6 2 C 4.9057453 2 4 2.9057453 4 4 L 4 20 C 4 21.094255 4.9057453 22 6 22 L 18 22 C 19.094255 22 20 21.094255 20 20 L 20 8 L 14 2 L 6 2 z M 6 4 L 13 4 L 13 9 L 18 9 L 18 20 L 6 20 L 6 4 z"/></svg>';
    });

  document.querySelectorAll("fluent-tree-view > fluent-tree-item[id*='node']").forEach((item) => {
    item.shadowRoot.querySelector("slot").innerHTML =
      '<svg xmlns="http://www.w3.org/2000/svg" fill="#000000" viewBox="0 0 24 24" width="24px" height="24px"><path d="M 4 4 C 2.9057453 4 2 4.9057453 2 6 L 2 18 C 2 19.094255 2.9057453 20 4 20 L 20 20 C 21.094255 20 22 19.094255 22 18 L 22 8 C 22 6.9057453 21.094255 6 20 6 L 12 6 L 10 4 L 4 4 z M 4 6 L 9.171875 6 L 11.171875 8 L 20 8 L 20 18 L 4 18 L 4 6 z"/></svg>';
  });
  document.getElementById("documents").style.display = "";
  document.getElementById("myInput").onkeyup = filter;
  document.getElementById("addDocument").onclick = addDocument;
  return false;
}

//filter by folder/filename
export function filter() {
  var input, filter, ul, li, a, i, txtValue;
  input = document.getElementById("myInput");
  filter = input.value.toUpperCase();
  ul = document.getElementById("documents");
  li = ul.getElementsByTagName("fluent-tree-item");
  for (i = 0; i < li.length; i++) {
    a = li[i];
    txtValue = a.textContent || a.innerText;
    if (txtValue.toUpperCase().indexOf(filter) > -1) {
      li[i].style.display = "";
    } else {
      li[i].style.display = "none";
    }
  }
  Array.from(document.getElementsByClassName("groupName")).forEach((element) => {
    element["style"].display = "none";
  });
}

//add selected documents
export function addDocument() {
  debugger;
  var allDocs = document.querySelectorAll("div#documentChooser fluent-checkbox");
  allDocs.forEach((doc) => {
    if (doc.getAttribute("current-checked") == "true") {
      var itemId = doc.parentElement.getAttribute("id");
      var itemTitle = doc.parentElement.innerText;
      doc.setAttribute("current-checked", "false");
      doc.setAttribute("disabled", "");
      var settings = {
        url:
          "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/download-document?id=" +
          itemId,
        method: "GET",
        timeout: 0,
        headers: {
          Authorization: "Bearer " + token,
          Cookie: "PHPSESSID=fffl40uonca198ah0opl7jt5un",
        },
      };

      $.ajax(settings).done(function (response) {
        var blob = new Blob([response], {
          // type: "application/pdf",
        });

        var reader = new FileReader();
        reader.readAsDataURL(blob);
        reader.onload = function () {
          var final = reader.result.toString().replace(/^data:.+;base64,/, "");
          Office.context.mailbox.item.addFileAttachmentFromBase64Async(final, itemTitle);
        };
      });
    }
  });
  return false;
}
