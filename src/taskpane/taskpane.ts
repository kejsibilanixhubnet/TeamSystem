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
//
import { creds } from "./creds";

//authorization token
var token = "";
var chosenPractice = "";
var indexOfPractice = 0;
var totalNumberOfPractices = 0;
var searchedForPractice = false;

//initialize
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("CloseForm").onclick = CloseForm;
    var ClientSecretKey = Office.context.roamingSettings.get("ClientSecretKey");
    if (ClientSecretKey != undefined) {
      document.getElementById("choosePracticeButton").onclick = searchForPractice;
      document.getElementById("nextPageOfPractices").onclick = goToNextPage;
      document.getElementById("prevPageOfPractices").onclick = goToPreviousPage;

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

//close form
export function CloseForm() {
  document.getElementById("myForm").style.display = "none";
}

//save token
export function SaveToken() {
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
    },
    error: function () {
      document.getElementById("SaveToken").onclick = SaveToken;

      document.getElementById("TokenRequired").style.display = "";
      document.getElementById("startingLoader").style.display = "none";
    },
  });
  Office.context.roamingSettings.saveAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {} else {
      document.getElementById("TokenRequired").style.display = "none";
      Office.onReady();
    }
  });
}

//get all practices
export async function GetAllPratiche(token) {
  var settings = {
    url: "https://howling-crypt-47129.herokuapp.com/https://testenv18.netlex.cloud/api-v2/archive",
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

        if (Office.context.mailbox.item.displayReplyForm != undefined) {
          //--------------------------------------------------------------------------------------------------------------------------------------------> Here here - received mail
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
    users.forEach((user) => {
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
    var userId = document
      .querySelector('fluent-option[aria-selected="true"]')
      .getAttribute("id")
      .replace("option-", "");
    var searchBy = document.getElementById("ricerca_testo")["value"];
    var paramString = "?ricerca_testo=" + searchBy + "&id_utente_soggetto=" + userId;
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
            chooseTemplateOrDocuments(); //----------------------------------------------------------------------------------------------------------> Here here - new mail
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
  document.getElementById("selectTemplate").onchange = addTemplate;
  return false;
}

//choose between template or documents
export async function chooseTemplateOrDocuments() {
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

//documents section
export async function chooseDocument(id) {
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

  document.getElementById("chooseTemplateOrDocs").style.display = "none";
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
