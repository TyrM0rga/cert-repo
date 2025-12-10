var global_host = parent.window.location.href;
var me_id = "";
var me_name = "";
var me_email = "";
var renewals = [];
var certs = [];
var cert_types = [];
var bestowed = [];
var admin = false;
let duplicate = false;
let file;
let usr_id;
let c_type;
var certificate = null;
var file_count = 0;
const notifyMonths = 3;

let up = ' <span class="arrow">&#9650;</span>';
let down = ' <span class="arrow">&#9660;</span>';

window.onerror = function (message, source, lineno) {
  alert(message + "\n\nResource: " + source + "\n\nLine #: " + lineno);
};

// add jquery dialog to newAwardDate function
$(document).ready(function () {
  // alert("test 0");
  var path = window.location.pathname;
  var page = "/" + path.split("/").pop();
  global_host = global_host.substring(0, global_host.indexOf(page));
  $("#loading").show();
  getMe();
  $("#certificateFile").on("change", function () {
    const files = this.files;
    handleCertificate(files);
  });
});

function dragOverHandler(e) {
  e.preventDefault();
  e.stopPropagation();
  $("#uploadArea").css({ "background-color": "#4a6cff" });
}

function dragOut() {
  $("#uploadArea").css({ "background-color": "#f9fafb" });
}

function dropHandler(e) {
  e.preventDefault();
  e.stopPropagation();
  $("#uploadArea").css({ "background-color": "#f9fafb" });
  let files = e.dataTransfer.files;
  handleCertificate(files);
  $("#submit").prop("disabled", false);
}

function browseFile(e) {
  e.preventDefault();
  e.stopPropagation();
  $("#certificateFile").click();
  $("#submit").prop("disabled", false);
}

function handleCertificate(files) {
  for (let i = 0; i < files.length; i++) {
    // alert(`File name: ${files[i].name}, file type: ${files[i].type}`);
    file_count++;
  }
  if (!files || files.length === 0) {
    return;
  }

  // destructuring assignment
  // sets certificate = files[0]
  [certificate] = files;

  // Reject multiple files
  if (files.length > 1) {
    showValidationDialog("Too many files!", "uploadArea");
    $("#certificateFile").val("");
    return;
  }

  if (certificate.type != "application/pdf") {
    showValidationDialog(
      "Document must be public document format (PDF).",
      "uploadArea"
    );
    return;
  }

  // file = e.dataTransfer.files;
  // for (let i = 0; i < file.length; i++) {
  //   file_count++;
  // }
  // certificate = file[0];
  // let cert = file[0];

  // let filename = file[0].name;

  $("#file-area").hide();
  let msg = `
  <h3>${certificate.name}</h3>
  `;
  $("#file-info").html(msg);
  $("#file-info").show();
}

function submitCertificate() {
  let isValid = true;

  // Check certificateType
  if (!$("#certificateType").val()) {
    showValidationDialog(
      "Please select a certificate type.",
      "certificateType"
    );
    isValid = false;
    return false; // Exit immediately
  }

  // Check recipient
  if (!$("#bestow_mem").val()) {
    showValidationDialog("Please select an award recipient.", "bestow_mem");
    isValid = false;
    return false; // Exit immediately
  }

  // Check award date
  if (!$("#awardDate").val()) {
    showValidationDialog("Please select an award date.", "awardDate");
    isValid = false;
    return false; // Exit immediately
  }

  // Check certificate file
  if (file_count < 1) {
    showValidationDialog(
      "Please upload an award certificate.",
      "certificateFile"
    );
    isValid = false;
    return false; // Exit immediately
  }

  if (isValid) {
    getDigest();
  }
}

// Helper function for showing the dialog
function showValidationDialog(message, fieldId) {
  $("#alert-msg").html(message);
  $("#dialog-message").dialog({
    dialogClass: "no-close",
    closeOnEscape: false,
    modal: true,
    title: "Missing Information",
    width: 500,
    position: { my: "center", at: "center", of: window },
    buttons: {
      Ok: function () {
        $(this).dialog("close");
        // Set focus to the empty field
        $("#" + fieldId).focus();
      },
    },
  });
}

function getMe() {
  if (sessionStorage.user_id) {
    me_id = parseInt(sessionStorage.user_id);
    me_name = sessionStorage.user_name;
    me_email = sessionStorage.user_email;
    if (sessionStorage.impersonate == "true") {
      var user_label =
        "<div style='position:absolute;top:2px;width:100%;color:#fff;font-weight:bold;font-size:12px;text-align:center;'>" +
        sessionStorage.user_name +
        "</div>";
      $("body").append(user_label);
    }
    checkAdmin();
  } else {
    var call = $.ajax({
      async: true,
      url: global_host + "/_api/web/CurrentUser",
      method: "GET",
      headers: { Accept: "application/json; odata=verbose" },
    });

    call.done(function (data) {
      me_id = data.d.Id;
      me_name = data.d.Title;
      me_email = data.d.Email;
      sessionStorage.user_id = me_id;
      sessionStorage.user_name = me_name;
      sessionStorage.user_email = me_email;
      checkAdmin();
    });

    call.fail(function (error) {
      displayError("Get Me", error);
    });
  }
}

function checkAdmin() {
  var queryUrl =
    global_host + "/_api/web/sitegroups/getbyname('CXR Owners')/users";

  var call = $.ajax({
    async: true,
    url: queryUrl,
    method: "GET",
    headers: {
      Accept: "application/json; odata=verbose",
    },
  });

  call.done(function (data) {
    for (var i = 0; i < data.d.results.length; i++) {
      var id = data.d.results[i].Id;
      if (id == me_id) {
        admin = true;
        break;
      }
    }

    if (admin) {
      $(".btn-container").show();
    } else {
      $(".btn-container").hide();
    }
    getCertTypes();
  });

  call.fail(function (error) {
    displayError("Check Admin ", error);
  });
}

// new rest call to certifications
// pop cert_types array
function getCertTypes() {
  var queryUrl =
    global_host +
    "/_api/Web/Lists/GetByTitle('Certifications')/Items?&$orderby=Sequence asc";

  var call = $.ajax({
    async: true,
    url: queryUrl,
    method: "GET",
    headers: {
      Accept: "application/json; odata=verbose",
    },
  });

  call.done((data) => {
    for (var i = 0; i < data.d.results.length; i++) {
      cert_types.push({
        id: data.d.results[i].Id,
        title: data.d.results[i].Title,
        long_title: data.d.results[i].LongTitle,
        sequence: data.d.results[i].Sequence,
      });
    }
    getCXRPersonnel();
  });

  call.fail(function (error) {
    displayError("Get Cert Types", error);
  });
}

function showAddCertDiv() {
  $("#cert_list").fadeOut("fast", function () {
    $("#add_cert").fadeIn("fast");
  });
}

function hideAddCertDiv() {
  $("#add_cert").fadeOut("fast", () => {
    $("#cert_list").fadeIn("fast");
  });
}

function getRenewals() {
  var queryUrl =
    global_host +
    "/_api/Web/Lists/GetByTitle('Renewals')/Items?" +
    "$select=Id,Title,RenewalDate,Certificate/Id,Certificate/Title,File/Name" +
    "&$expand=Certificate,File" +
    "&$orderby=RenewalDate desc" +
    "&$top=5000";

  var call = $.ajax({
    async: true,
    url: queryUrl,
    method: "GET",
    headers: {
      Accept: "application/json; odata=verbose",
    },
  });

  call.done((data) => {
    data.d.results.forEach((item) => {
      renewals.push({
        id: item.Id,
        title: item.Title,
        renewal_date: item.RenewalDate,
        file_name: item.File.Name,
        cert_id: item.Certificate.Id,
      });
    });
  });

  call.fail(function (error) {
    displayError("Get Certificates", error);
  });
}

function getCertificates() {
  var queryUrl =
    global_host +
    "/_api/Web/Lists/GetByTitle('Certificates')/Items?" +
    "$select=Id,Title,Certification/Id,Certification/Title,Certification/Sequence,AwardDate,File/Name," +
    "BestowMember/Id,BestowMember/Title" +
    "&$expand=Certification,BestowMember,File" +
    "&$orderby=BestowMember/Title asc, Certification/Sequence asc" +
    "&$top=5000";

  var call = $.ajax({
    async: true,
    url: queryUrl,
    method: "GET",
    headers: {
      Accept: "application/json; odata=verbose",
    },
  });

  call.done((data) => {
    for (var i = 0; i < data.d.results.length; i++) {
      certs.push({
        id: data.d.results[i].Id,
        title: data.d.results[i].Title,
        name: data.d.results[i].File.Name,
        Certification_id: data.d.results[i].Certification.Id,
        Certification_Title: data.d.results[i].Certification.Title,
        certification_seq: data.d.results[i].Certification.Sequence,
        date: data.d.results[i].AwardDate,
        bestowMem_id: data.d.results[i].BestowMember.Id,
        bestowMem_title: data.d.results[i].BestowMember.Title,
      });
    }
    certs.forEach((cert) => {
      const match_id = renewals.find((renewal) => cert.id == renewal.cert_id);

      if (match_id) {
        cert.date = match_id.renewal_date;
      }
    });
    // alert(typeof certs[1].certification_seq);
    buildTable(certs);
    let mydate = new Date(certs[0].date);
    mydate.toLocaleDateString();
    // alert(mydate);
  });

  call.fail(function (error) {
    displayError("Get Certificates", error);
  });
}

function buildTable(cert_arr) {
  let html = "";
  for (let i = 0; i < cert_arr.length; i++) {
    let awd_date = new Date(cert_arr[i].date);
    // formatted award date
    let fawd_date = awd_date.toLocaleDateString();

    // set renew date 5 years from award date
    let renew_date = "";
    if (cert_arr[i].Certification_Title === "AFCEM") {
      let renew_dt = new Date(fawd_date);
      renew_dt.setFullYear(renew_dt.getFullYear() + 5);
      renew_date = renew_dt.toLocaleDateString();
    }

    // awd_date = awd_date.toLocaleDateString();
    let name = cert_arr[i].name;
    if (admin || cert_arr[i].bestowMem_id == me_id) {
      html += `
      <tr id="${cert_arr[i].id}" onclick="openDocument('${name}')" class="row_pointer">
        <td>${cert_arr[i].bestowMem_title}</td>
        <td>${cert_arr[i].Certification_Title}</td>
        <td>${fawd_date}</td>
        <td class="renew-dt" data-award-date="${cert_arr[i].date}">${renew_date}</td>        
        <td><span class="material-symbols-outlined">download</span></td>
      </tr>`;
    } else {
      html += `
      <tr id="${cert_arr[i].id}">
        <td>${cert_arr[i].bestowMem_title}</td>
        <td>${cert_arr[i].Certification_Title}</td>
        <td>${fawd_date}</td>
        <td class="renew-dt" data-award-date="${cert_arr[i].date}">${renew_date}</td>  
        <td></td>
      </tr>`;
    }
  }

  $("tbody")
    .html(html)
    .promise()
    .done(function () {
      highlightRenewalDates();
      $("#blanket").fadeOut("fast", function () {
        $("#page").fadeIn("fast");
      });
    });
}

function applyFilter() {
  let searchTerm = $("#searchInput").val();
  searchTerm = searchTerm.toLowerCase();
  // filteredData = certs;
  filteredData = certs.filter((cert) =>
    cert.bestowMem_title.toLowerCase().includes(searchTerm)
  );
  buildTable(filteredData);
}

function clearSearchFilter(){
  $("#searchInput").val('')
  buildTable(certs)
}

function sortRecipients() {
  $("thead th .arrow").remove();
  const isSorted = isArrSorted(certs, "bestowMem_title");
  if (isSorted) {
    $("#member").append(up);
    certs.sort((a, b) => b.bestowMem_title.localeCompare(a.bestowMem_title));
  } else {
    $("#member").append(down);
    certs.sort((a, b) => a.bestowMem_title.localeCompare(b.bestowMem_title));
  }
  buildTable(certs);
}

function isArrSorted(arr, key) {
  for (let i = 0; i < arr.length - 1; i++) {
    if (arr[i][key].localeCompare(arr[i + 1][key]) > 0) {
      return false;
    }
  }
  return true;
}

function sortCertification() {
  $("thead th .arrow").remove();
  const isSorted = isSeqSorted(certs, "certification_seq");
  if (isSorted) {
    $("#award").append(up);
    certs.sort((a, b) => {
      const certA = parseInt(a.certification_seq);
      const certB = parseInt(b.certification_seq);
      return certB - certA;
    });
  } else {
    $("#award").append(down);
    certs.sort((a, b) => {
      const certA = parseInt(a.certification_seq);
      const certB = parseInt(b.certification_seq);
      return certA - certB;
    });
  }
  buildTable(certs);
}

function isSeqSorted(arr, key) {
  for (let i = 0; i < arr.length - 1; i++) {
    const seqA = parseInt(arr[i][key]);
    const seqB = parseInt(arr[i + 1][key]);
    if (seqA > seqB) {
      return false;
    }
  }
  return true;
}

function sortDate() {
  $("thead th .arrow").remove();
  const isSorted = isDateSorted(certs, "date");
  if (isSorted) {
    $("#date").append(up);
    certs.sort((a, b) => new Date(b.date) - new Date(a.date));
  } else {
    $("#date").append(down);
    certs.sort((a, b) => new Date(a.date) - new Date(b.date));
  }
  buildTable(certs);
}

function isDateSorted(arr, key) {
  for (let i = 0; i < arr.length - 1; i++) {
    let dateA = new Date(arr[i][key]);
    let dateB = new Date(arr[i + 1][key]);
    if (dateA > dateB) {
      return false;
    }
  }
  return true;
}

function highlightRenewalDates() {
  // Get current date and date 6 months from now
  const today = new Date();
  const sixMonthsFromNow = new Date(today);
  sixMonthsFromNow.setMonth(today.getMonth() + notifyMonths);

  $(".renew-dt").each(function () {
    if ($(this).text() !== "") {
      let renewDateText = $(this).text();

      let renewDate = new Date(renewDateText);

      // let renewalDate = new Date($(this).text());

      // Check if renewal date is within next 6 months
      if (renewDate >= today && renewDate <= sixMonthsFromNow) {
        $(this).css({
          color: "#b89000",
          "font-weight": "bold",
        });
      } else if (renewDate < today) {
        // Red for expired renewals
        $(this).css({ color: "#FF0000", "font-weight": "bold" });
      }
    }
  });
}

function openDocument(file_name) {
  var path = global_host + "/Certificates/" + file_name;
  window.open(path);
}

function getCXRPersonnel() {
  var queryUrl =
    global_host +
    "/_api/web/sitegroups/getbyname('CXR Members')/users?&$top=5000";
  var call = $.ajax({
    async: true,
    url: queryUrl,
    method: "GET",
    headers: {
      Accept: "application/json; odata=verbose",
    },
  });

  call.done(function (data) {
    // hide spinner
    $("#loading").hide();
    // Update dropdown placeholder
    $("#bestow_mem option:first").text("Select a person");
    for (var i = 0; i < data.d.results.length; i++) {
      bestowed.push({
        id: data.d.results[i].Id,
        name: data.d.results[i].Title,
        email: data.d.results[i].Email,
      });
    }
    getEMPersonnel();
  });

  call.fail(function (error) {
    displayError("Get Award Recipient:", error);
  });
}

function getEMPersonnel() {
  var queryUrl =
    global_host +
    "/_api/web/sitegroups/getbyname('EM Personnel')/users?&$top=5000";
  var call = $.ajax({
    async: true,
    url: queryUrl,
    method: "GET",
    headers: {
      Accept: "application/json; odata=verbose",
    },
  });

  call.done(function (data) {
    // hide spinner
    $("#loading").hide();
    // Update dropdown placeholder
    // Need to add user-id -> mbr_id
    $("#bestow_mem option:first").text("Select a person");
    for (var i = 0; i < data.d.results.length; i++) {
      var title = data.d.results[i].Title;
      bestowed.push({
        id: data.d.results[i].Id,
        name: data.d.results[i].Title,
        email: data.d.results[i].Email,
      });
    }
    bestowed.sort((a, b) => a.name.localeCompare(b.name));
    // bestowed = bestowed.sort(sortName);
    setCertificateType();
  });

  call.fail(function (error) {
    displayError("Get Award Recipient:", error);
  });
}

function sortName(a, b) {
  if (a[1] < b[1]) return -1;
  if (a[1] > b[1]) return 1;
  return 0;
}

function setCertificateType() {
  let html = "";
  for (var i = 0; i < cert_types.length; i++) {
    html += `<option value="${cert_types[i].sequence}">${cert_types[i].long_title}</option>`;
  }
  $("#certificateType").append(html);
  setAwardRecipient();
}

function setAwardRecipient() {
  let html = "";
  for (var i = 0; i < bestowed.length; i++) {
    html = `<option value="${bestowed[i].id}">${bestowed[i].name}</option>`;
    $("#bestow_mem").append(html);
  }
  getCertificates();
  getRenewals();
}
// triggered via bestow_mem onchange event
function selectCertType() {
  $("#certificateType").prop("disabled", false);
  usr_id = $("#bestow_mem").val();
}

function newAwardDate() {
  c_type = $("#certificateType").val();

  duplicate =
    certs.find(
      (cert) => cert.bestowMem_id == usr_id && cert.certification_seq == c_type
    ) !== undefined;

  let cert_name = cert_types.find((dup) => dup.sequence == c_type);

  let title = cert_name.title;
  let mem_title = $("#bestow_mem option:selected").text();
  mem_title = mem_title.substring(0, mem_title.indexOf(" USAF"));
  let dup_message = `${title} certificate already exists for ${mem_title}.`;

  if (duplicate) {
    $("#alert-msg").html(dup_message);
    $("#dialog-message").dialog({
      dialogClass: "no-close",
      closeOnEscape: false,
      modal: true,
      title: "Duplicate Certificate",
      width: 500,
      position: { my: "center", at: "center", of: window },
      buttons: {
        Ok: function () {
          $(this).dialog("close");
          // Set focus to the empty field
          $("#certificateType").val("");
          $("#certificateType").focus();
        },
      },
    });
  } else {
    $("#awardDate").prop("disabled", false);
  }

  $("#awardDate").datepicker({
    dateFormat: "mm/dd/yy",
    changeMonth: true,
    changeYear: true,
    showButtonPanel: true,
    defaultDate: new Date(),
    yearRange: "-50:+0",
    maxDate: new Date(),
  });
}

function getDigest() {
  var call = $.ajax({
    async: true,
    url: global_host + "/_api/contextinfo",
    method: "POST",
    headers: {
      Accept: "application/json; odata=verbose",
    },
  });

  call.done(function (data) {
    digest = data.d.GetContextWebInformation.FormDigestValue;
    uploadDocument();
  });

  call.fail(function (error) {
    displayError("Get Digest", error);
  });
}

function uploadDocument() {
  let f_name;
  let ele_val = $("#certificateType").val();
  usr_id = $("#bestow_mem").val();
  const sel_certType = cert_types.find((cert) => cert.sequence == ele_val);
  let label = sel_certType.title;
  f_name = `${label}_${usr_id}.pdf`;

  let filename = certificate.name;
  //  [cert_type]_[user_id].pdf
  // filename = `${c_type}_${usr_id}.pdf`;
  var reader = new FileReader();
  reader.onload = (e) => {
    uploadFile(e.target.result, f_name);
  };
  reader.onerror = (e) => {
    alert(e.target.error);
  };

  reader.readAsArrayBuffer(certificate);

  function uploadFile(buffer, fileName) {
    var url = `${global_host}/_api/web/GetFolderByServerRelativeUrl('Certificates')/Files/Add(url='${fileName}',overwrite=true)`;

    var call = $.ajax({
      url: url,
      method: "POST",
      contentType: "application/json;odata=verbose",
      data: buffer,
      processData: false,
      headers: {
        Accept: "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "Content-Length": buffer.byteLength,
      },
    });

    call.done((data) => {
      getListItem(data.d);
    });

    call.fail((error) => {
      displayError("Upload Document - Upload File ", error);
    });
  }

  function getListItem(file) {
    var call = $.ajax({
      url: file.ListItemAllFields.__deferred.uri,
      type: "GET",
      dataType: "json",
      headers: {
        Accept: "application/json;odata=verbose",
      },
    });

    call.done(function (data) {
      updateListItemFields(data.d.Id);
    });

    call.fail(function (error) {
      displayError("Upload Document - Get List Item", error);
    });
  }

  function updateListItemFields(item_id) {
    let nomen = $("#bestow_mem option:selected").text();
    const cert_val = $("#certificateType").val();
    const cert_types = {
      1: "AFAHR",
      2: "AFAEM",
      3: "AFCEM",
    };

    let cert_type = cert_types[cert_val];
    var itemType = "SP.Data.CertificatesItem";
    var awardDate = new Date($("#awardDate").val());
    // var list_url = `${global_host}/_api/Web/Lists/getByTitle('Certificates')/Items('${item_id}')`
    var call = $.ajax({
      url:
        global_host +
        "/_api/Web/Lists/getByTitle('Certificates')/Items(" +
        item_id +
        ")",
      type: "POST",
      data: JSON.stringify({
        __metadata: { type: itemType },
        CertificationId: $("#certificateType").val(),
        BestowMemberId: $("#bestow_mem").val(),
        AwardDate: awardDate,
      }),
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      },
    });
    call.done((data) => {
      certs.push({
        id: certs.length + 1,
        title: null,
        name: certificate.name,
        Certification_id: $("#certificateType").val(),
        Certification_Title: cert_type,
        certification_seq: $("#certificateType").val(),
        date: awardDate,
        bestowMem_id: $("#bestow_mem").val(),
        bestowMem_title: nomen,
      });
      certs.sort((a, b) => a.bestowMem_title.localeCompare(b.bestowMem_title));
      buildTable(certs);
      $("#certificateType").val("");
      $("#awardDate").val("");
      $("#bestow_mem").val("");
      $("#fileName").empty();
      $("#file-info").hide();
      $("#file-area").show();
      $("#add_cert").fadeOut("fast", function () {
        $("#cert_list").fadeIn("fast");
      });
    });
  }
}

function resetFormFields() {
  $("#certificateType").val("");
}

function displayError(funcName, error) {
  console.log(funcName + ": " + JSON.stringify(error));
  $("#blanket").fadeOut("fast");
  var response = JSON.parse(error.responseText);
  var message = response ? response.error.message.value : textStatus;
  var page = location.pathname;
  for (var cnt = 0; cnt <= 0; cnt++) {
    var pos = page.lastIndexOf("/");
    page = page.substr(pos + 1);
  }
  $("#error_page").html("<b>Page:</b> " + page);
  $("#error_function").html("<b>Function Name:</b> " + funcName);
  $("#error_msg").html("<b>Error Message:</b> " + message);

  $(function () {
    $("#error-message").dialog({
      dialogClass: "no-close",
      modal: true,
      width: 600,
      position: { my: "center", at: "center", of: $("#page") },
      buttons: {
        Ok: function () {
          $(this).dialog("close");
        },
      },
    });
  });
}
