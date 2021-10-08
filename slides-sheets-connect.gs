/**
 * Copyright (c) 2021 Tably Inc.
 * Released under the MIT license
 */

const TOP_ROW = 2;
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
    const includeTimeInfo = form.includeTimeInfo;
    doCopySpeakerNotesFromSheetToSlide(url, includeTimeInfo === "yes");
  }
};

const promptToGetURL = (mode) => {
  const additional =
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
      : `
    <div style="display: flex; flex-direction: row; margin-bottom: 16px;">
      <label>Include Time Info</label>
      <input type="radio" name="includeTimeInfo" id="timeInfoYes" value="yes" />
      <label for="timeInfoYes">Yes</label>
      <input type="radio" name="includeTimeInfo" id="timeInfoNo" value="no" checked />
      <label for="timeInfoNo">No</label>
    </div>
  `;
  const html = `
    <form style="display: flex; flex-direction: column;" id="form">
      <div style="display: flex; flex-direction: row; margin-bottom: 16px;">
        <label for="slideUrl" style="margin-right: 8px;">Slide URL:</label>
        <input type="text" id="slideUrl" name="slideUrl" value="https://docs.google.com/presentation/d/1-0u_qFjZsiVsvWXLH2zaOoFAMbNUH3xhtTRJxSHRqVQ/edit#slide=id.gc990316775_0_0" />
      </div>
      ${additional}
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

const findTitleShape = (slide) => {
  const elements = slide.getPageElements();
  const shapes = [];
  let lowestTop = 1024;
  let shapeIndex = 0;
  let shapeIndexForTitle = -1;
  elements.forEach(function (element) {
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
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
  if (shapeIndexForTitle !== -1) {
    return shapes[shapeIndexForTitle];
  } else {
    return null;
  }
};

const importFromSlide = (url) => {
  const importFromSlideResult = [];
  const preso = SlidesApp.openByUrl(url); // Get Slide by URL
  const presentationId = preso.getId();
  const slides = preso.getSlides(); // Get all slides

  slides.forEach(function (slide, slideIndex) {
    const speakerNote = slide
      .getNotesPage()
      .getSpeakerNotesShape()
      .getText()
      .asString();
    const titleShape = findTitleShape(slide);
    const slideId = slide.getObjectId();
    const exportUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}/pages/${slideId}/thumbnail`;
    if (titleShape) {
      importFromSlideResult.push([
        titleShape.getText().asString(),
        speakerNote,
        exportUrl,
      ]);
    } else {
      importFromSlideResult.push(["", speakerNote, exportUrl]);
    }
  });

  return importFromSlideResult;
};

const pasteSlideTitlesAndNotesToSheet = (importFromSlideResult) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  sheet.getRange(1, 1).setValue("Slide #");
  sheet.getRange(1, 2).setValue("Title");
  sheet.getRange(1, 3).setValue("Speaker Note");
  sheet.getRange(1, 4).setValue("Thumbnail Image");
  sheet.getRange(1, 5).setValue("Duration");
  sheet.getRange(1, 6).setValue("Start");

  for (let i = 0; i < importFromSlideResult.length; i++) {
    const title = importFromSlideResult[i][0];
    let note = importFromSlideResult[i][1];
    let timeInfo = {
      duration: "00:00:00",
    };
    const m = note.match(/^\{.+\}/);
    if (m) {
      try {
        const timeInfoStr = m[0];
        timeInfo = JSON.parse(
          timeInfoStr.replaceAll("”", '"').replaceAll("“", '"')
        );
        note = note.substring(timeInfoStr.length);
      } catch (e) {
        console.error(e);
      }
    }
    sheet.getRange(TOP_ROW + i, TOP_COLUMN).setValue(i + 1);
    sheet.getRange(TOP_ROW + i, TOP_COLUMN + 1).setValue(title);
    sheet.getRange(TOP_ROW + i, TOP_COLUMN + 2).setValue(note);
    sheet
      .getRange(TOP_ROW + i, TOP_COLUMN + 4)
      .setNumberFormat("h:mm")
      .setValue(timeInfo.duration);
    sheet
      .getRange(TOP_ROW + i, TOP_COLUMN + 5)
      .setNumberFormat("h:mm")
      .setValue(
        i === 0
          ? 0
          : "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))+INDIRECT(ADDRESS(ROW()-1,COLUMN()-1))"
      );
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
    sheet.setRowHeights(TOP_ROW, rowCount, 90);
    sheet.setColumnWidth(4, 160);
  } else {
    sheet.autoResizeRows(TOP_ROW, rowCount);
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const importFromSlideResult = [];
  for (let i = 0; i < sheet.getLastRow(); i++) {
    const duration = sheet.getRange(TOP_ROW + i, TOP_COLUMN + 4).getValue();
    const start = sheet.getRange(TOP_ROW + i, TOP_COLUMN + 5).getValue();
    importFromSlideResult.push({
      title: sheet.getRange(TOP_ROW + i, TOP_COLUMN + 1).getValue(),
      note: sheet.getRange(TOP_ROW + i, TOP_COLUMN + 2).getValue(),
      duration: duration
        ? Utilities.formatDate(duration, "JST", "HH:mm:ss")
        : undefined,
      start: start ? Utilities.formatDate(start, "JST", "HH:mm:ss") : undefined,
    });
  }
  return importFromSlideResult;
};

const pasteSlideTitlesAndNotesToSlide = (
  url,
  importFromSlideResult,
  includeTimeInfo
) => {
  const preso = SlidesApp.openByUrl(url); // Get Slide by Opening URL
  const slides = preso.getSlides(); // Get all slides

  slides.forEach(function (slide, slideIndex) {
    const shapes = slide.getShapes();

    const noteTextRange = slide.getNotesPage().getSpeakerNotesShape().getText();
    if (importFromSlideResult[slideIndex].note == "") {
      noteTextRange.setText("");
    } else {
      const duration = importFromSlideResult[slideIndex].duration;
      const start = importFromSlideResult[slideIndex].start;
      if (includeTimeInfo && duration && start) {
        const timeInfo = {
          duration,
          start,
        };
        noteTextRange.setText(
          `${JSON.stringify(timeInfo)} ${
            importFromSlideResult[slideIndex].note
          }`
        );
      } else {
        noteTextRange.setText(importFromSlideResult[slideIndex].note);
      }
    }

    const titleShape = findTitleShape(slide);
    if (titleShape) {
      titleShape.getText().setText(importFromSlideResult[slideIndex].title);
    }
  });
};

const copySpeakerNotesFromSheetToSlide = () => {
  promptToGetURL("copySpeakerNotesFromSheetToSlide");
};

const doCopySpeakerNotesFromSheetToSlide = (url, includeTimeInfo) => {
  const importFromSlideResult = importFromSheet();
  console.log(importFromSlideResult);
  pasteSlideTitlesAndNotesToSlide(url, importFromSlideResult, includeTimeInfo);
};
