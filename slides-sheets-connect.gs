const TOP_ROW = 1;
const TOP_COLUMN = 1;

const onOpen = () => {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Slides ")
    .addItem("Copy from Slides", "copyFromSlideToSheet")
    .addItem("Write to Slides", "copySpeakerNotesFromSheetToSlide")
    .addToUi();
};

const onSubmitForm = (form) => {
  const url = form.slideUrl;
  if (!url) {
    return;
  }
  const mode = form.mode;
  if (mode === "copyFromSlideToSheet") {
    const importImages = form.importImages;
    doCopyFromSlideToSheet(url, importImages === "yes");
  } else if (mode === "copySpeakerNotesFromSheetToSlide") {
    doCopySpeakerNotesFromSheetToSlide(url);
  }
};

const promptToGetURL = (mode) => {
  const importImages =
    mode === "copyFromSlideToSheet"
      ? `
    <div style="display: flex; flex-direction: row; margin-bottom: 16px;">
      <label>Import Images</label>
      <input type="radio" name="importImages" id="importImagesYes" value="yes" />
      <label for="importImagesYes">Yes</label>
      <input type="radio" name="importImages" id="importImagesNo" value="no" checked />
      <label for="importImagesNo">No</label>
    </div>
  `
      : "";
  const html = `
    <form style="display: flex; flex-direction: column;" id="form">
      <div style="display: flex; flex-direction: row; margin-bottom: 16px;">
        <label for="slideUrl" style="margin-right: 8px;">Slide URL:</label>
        <input type="text" id="slideUrl" name="slideUrl" />
      </div>
      ${importImages}
      <div style="display: flex; flex-direction: row; justify-content: flex-end;">
        <button type="button" onclick="google.script.run.onSubmitForm(document.querySelector('#form')); google.script.host.close();">OK</button>
        <button type="button" onclick="google.script.host.close();" style="margin-left: 16px;">Cancel</button>
      </div>
      <input type="hidden" name="mode" value="${mode}" />
    </form>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Target Slides URL");
};

const importFromSlide = (url) => {
  let importFromSlideResult = [];
  let preso = SlidesApp.openByUrl(url); // Get Slide by URL
  const presentationId = preso.getId();
  let slides = preso.getSlides(); // Get all slides

  slides.forEach(function (slide, slideIndex) {
    // Process each slide
    console.log(`slideIndex: ${slideIndex}`);

    let shapeIndexForTitle = 0;
    let lowestTop = 1024; // Not sure about the highest vertical point in slides

    let speakerNote = slide
      .getNotesPage()
      .getSpeakerNotesShape()
      .getText()
      .asString();

    const elements = slide.getPageElements();
    const shapes = [];
    let shapeIndex = 0;
    elements.forEach(function (element) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        if (shape.getShapeType() == SlidesApp.ShapeType.TEXT_BOX) {
          let shapeTop = shape.getTop();
          if (shapeTop < lowestTop) {
            lowestTop = shapeTop;
            shapeIndexForTitle = shapeIndex;
          }
        }
        shapes.push(shape);
        shapeIndex++;
      }
    });

    const slideId = slide.getObjectId();
    const exportUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}/pages/${slideId}/thumbnail`;
    if (shapes.length != 0) {
      if (
        shapes[shapeIndexForTitle].getShapeType() ==
        SlidesApp.ShapeType.TEXT_BOX
      ) {
        importFromSlideResult[slideIndex] = [
          shapes[shapeIndexForTitle].getText().asString(),
          speakerNote,
          exportUrl,
        ];
      } else {
        importFromSlideResult[slideIndex] = ["", speakerNote, exportUrl];
      }
    } else {
      importFromSlideResult[slideIndex] = ["", speakerNote, exportUrl];
    }
  });

  return importFromSlideResult;
};

const pasteSlideTitlesAndNotesToSheet = (importFromSlideResult) => {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  for (let i = 0; i < importFromSlideResult.length; i++) {
    sheet.getRange(TOP_ROW + i, TOP_COLUMN).setValue(i + 1);
    sheet
      .getRange(TOP_ROW + i, TOP_COLUMN + 1)
      .setValue(importFromSlideResult[i][0]);
    sheet
      .getRange(TOP_ROW + i, TOP_COLUMN + 2)
      .setValue(importFromSlideResult[i][1]);
  }
};

const fetchSlideImagesAndPasteThemToSheet = (importFromSlideResult) => {
  const token = ScriptApp.getOAuthToken();
  for (let i = 0; i < importFromSlideResult.length; i++) {
    const exportUrl = importFromSlideResult[i][2];
    let response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        followRedirects: true,
      },
    });
    const responseData = JSON.parse(response.getContentText());
    const contentUrl = responseData.contentUrl;
    response = UrlFetchApp.fetch(contentUrl, {
      followRedirects: true,
    });
    const blob = response.getBlob();
    const dataUrl = `data:image/png;base64,${Utilities.base64Encode(
      blob.getBytes()
    )}`;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getRange(TOP_ROW + i, TOP_COLUMN + 3);
    range.setValue('=IMAGE("http")');
    const builder = range.getValue().toBuilder();
    builder.setSourceUrl(dataUrl);
    const cellImage = builder.build();
    range.setValue(cellImage);
  }
};

const adjustRowHeight = (rowCount, importImages) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (importImages) {
    sheet.setRowHeights(1, rowCount, 90);
    sheet.setColumnWidth(4, 160);
  } else {
    sheet.autoResizeRows(1, rowCount);
  }
};

const copyFromSlideToSheet = () => {
  promptToGetURL("copyFromSlideToSheet");
};

const doCopyFromSlideToSheet = (url, importImages) => {
  let importFromSlideResult = importFromSlide(url);
  pasteSlideTitlesAndNotesToSheet(importFromSlideResult);
  if (importImages) {
    fetchSlideImagesAndPasteThemToSheet(importFromSlideResult);
  }
  adjustRowHeight(importFromSlideResult.length, importImages);
};

const importFromSheet = () => {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Initialize slideTitleAndNotes Array
  let importFromSlideResult = [];
  for (let i = 0; i < sheet.getLastRow(); i++) {
    importFromSlideResult[i] = [];
  }

  for (let i = 0; i < sheet.getLastRow(); i++) {
    importFromSlideResult[i][0] = sheet
      .getRange(TOP_ROW + i, TOP_COLUMN + 1)
      .getValue();
    importFromSlideResult[i][1] = sheet
      .getRange(TOP_ROW + i, TOP_COLUMN + 2)
      .getValue();
  }

  return importFromSlideResult;
};

const pasteSlideTitlesAndNotesToSlide = (url, importFromSlideResult) => {
  let preso = SlidesApp.openByUrl(url); // Get Slide by Opening URL
  let slides = preso.getSlides(); // Get all slides

  slides.forEach(function (slide, slideIndex) {
    let shapes = slide.getShapes();
    let shapeIndexForTitle = 0;
    let lowestTop = 1024; // Not sure about the highest vertical point in slides

    if (importFromSlideResult[slideIndex][1] == "") {
      slide.getNotesPage().getSpeakerNotesShape().getText().setText("");
    } else {
      slide
        .getNotesPage()
        .getSpeakerNotesShape()
        .getText()
        .setText(importFromSlideResult[slideIndex][1]);
    }

    shapes.forEach(function (shape, index) {
      if (shape.getShapeType == SlidesApp.ShapeType.TEXT_BOX) {
        let shapeTop = shape.getTop();
        // Logger.log(shapeTop); // FOR DEBUG
        if (shapeTop < lowestTop) {
          lowestTop = shapeTop;
          shapeIndexForTitle = index;
        }
      }
    });

    if (shapes.length != 0) {
      if (
        shapes[shapeIndexForTitle].getShapeType() ==
        SlidesApp.ShapeType.TEXT_BOX
      ) {
        shapes[shapeIndexForTitle]
          .getText()
          .setText(importFromSlideResult[slideIndex][0]);
      }
    }
  });
};

const copySpeakerNotesFromSheetToSlide = () => {
  promptToGetURL("copySpeakerNotesFromSheetToSlide");
};

const doCopySpeakerNotesFromSheetToSlide = (url) => {
  let importFromSlideResult = importFromSheet();
  pasteSlideTitlesAndNotesToSlide(url, importFromSlideResult);
};
