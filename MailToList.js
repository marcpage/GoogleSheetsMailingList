function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Actions')
      .addItem('Send Emails', 'send_emails')
      .addToUi();
}

function get_draft_by_name(draft_name) {
  var drafts = GmailApp.getDraftMessages();
  
  for (let draftIndex = 0; draftIndex < drafts.length; ++draftIndex) {
    
    if (drafts[draftIndex].getSubject() == draft_name) {
      return drafts[draftIndex];
    }

  }
  
  return undefined;
}

function match_attachments_from_body(inlineImages, body) {
  var name_image = {};
  var inlinePattern = /<img[^>]+data-surl="cid:([^"]+)" src="cid:([^"]+)" alt="([^"]+)"[^>]+>/g;
  var inlinedImageInfo = [...body.matchAll(inlinePattern)];

  for (var inlinedIndex = 0; inlinedIndex < inlinedImageInfo.length; ++inlinedIndex) {
    if (inlinedImageInfo[inlinedIndex][1] != inlinedImageInfo[inlinedIndex][2]) {
      throw "Mismatched identifiers";
    }
  }

  if (inlinedImageInfo.length != inlineImages.length) {
    throw "Mismatched inlined image counts";
  }

  for (var imagesIndex = 0; imagesIndex < inlineImages.length; ++imagesIndex) {
    var image = inlineImages[imagesIndex];
    var imageInfo = inlinedImageInfo[imagesIndex];

    if (imageInfo[3] != image.getName()) {
      throw "Image name mismatch: " + imageInfo[3] + " vs " + image.getName();
    }

    name_image[imageInfo[1]] = image;
  }

  return name_image;
}

function send_message(draft, name, recipient) {
  var inlineImages = draft.getAttachments({"includeAttachments": false, "includeInlineImages": true});
  var attachments = draft.getAttachments({"includeAttachments": true, "includeInlineImages": false});
  var plainBody = draft.getPlainBody();
  var htmlBody = draft.getBody();

  console.log("Sending email " + draft.getSubject() + " to " + recipient);
  GmailApp.sendEmail(
      recipient,
      draft.getSubject(), 
      name + ",\r\n" + plainBody,
      {
        "htmlBody": name + ",<br/>" + htmlBody,
        "attachments": attachments,
        "cc": draft.getCc(),
        "bcc": draft.getBcc(),
        "inlineImages": match_attachments_from_body(inlineImages, draft.getBody())
      });

}

function send_emails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var contents = sheet.getSheetValues(1,1, sheet.getDataRange().getHeight(), 3);
  var draft = get_draft_by_name(sheet.getName());

  if (!draft) {
    SpreadsheetApp.getUi().alert('We were not able to find a draft email named "' + sheet.getName() + '"');
  } else {

    for (let row_index = 1; row_index < contents.length; ++row_index) {
      var name = contents[row_index][0];
      var email = contents[row_index][1];
      var sent = contents[row_index].length > 2 ? contents[row_index][2] == "Sent" : false;

      console.log(name + " <" + email + "> already sent? " + sent);
      if (!sent) {
        send_message(draft, name, email);
      }
      sheet.getRange(row_index + 1, 3).setValue("Sent");
    }

  }

}

